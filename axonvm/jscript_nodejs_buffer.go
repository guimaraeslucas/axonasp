/*
 * AxonASP Server
 * Copyright (C) 2026 G3pix Ltda. All rights reserved.
 *
 * Developed by Lucas Guimarães - G3pix Ltda
 * Contact: https://g3pix.com.br
 * Project URL: https://g3pix.com.br/axonasp
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *
 * Attribution Notice:
 * If this software is used in other projects, the name "AxonASP Server"
 * must be cited in the documentation or "About" section.
 *
 * Contribution Policy:
 * Modifications to the core source code of AxonASP Server must be
 * made available under this same license terms.
 */
package axonvm

import (
	"encoding/base64"
	"encoding/hex"
	"fmt"
	"strings"
)

// jsBuffer represents a Node.js-compatible Buffer object wrapping a Go []byte.
// Stored in vm.jsBufferItems[id] for lifecycle management.
type jsBuffer struct {
	data []byte // Underlying byte slice
}

// jsCreateBufferConstructor creates the Buffer global constructor function.
// Buffer is implemented as a JScript object with static methods and a __js_ctor marker.
func (vm *VM) jsCreateBufferConstructor() Value {
	objID := vm.allocJSID()
	obj := make(map[string]Value, 10)

	// Type and constructor markers
	obj["__js_type"] = NewString("Function")
	obj["__js_ctor"] = NewString("Buffer")
	obj["name"] = NewString("Buffer")
	obj["length"] = NewInteger(1)

	// Static methods: Buffer.from(), Buffer.alloc(), Buffer.isBuffer()
	// These are stored as special markers; dispatch happens in jsCallMember
	obj["__js_from_method"] = NewString("__js_buffer_from__")
	obj["__js_alloc_method"] = NewString("__js_buffer_alloc__")
	obj["__js_isbuffer_method"] = NewString("__js_buffer_isbuffer__")

	vm.jsObjectItems[objID] = obj
	vm.jsPropertyItems[objID] = make(map[string]jsPropertyDescriptor, 10)

	return Value{Type: VTJSFunction, Num: objID}
}

// jsCreateBufferInstance creates a new Buffer instance wrapping the given byte slice.
// The instance stores the byte data and provides methods for conversion and access.
func (vm *VM) jsCreateBufferInstance(data []byte) Value {
	bufID := vm.allocJSID()
	obj := make(map[string]Value)

	// Type marker and reference to underlying data
	obj["__js_type"] = NewString("Buffer")
	obj["__js_buffer_data"] = NewString(fmt.Sprintf("__buffer_%d", bufID)) // Reference key

	// Store the actual buffer data
	if vm.jsBufferItems == nil {
		vm.jsBufferItems = make(map[int64]*jsBuffer)
	}
	vm.jsBufferItems[bufID] = &jsBuffer{data: data}

	// Properties
	obj["length"] = NewInteger(int64(len(data)))

	vm.jsObjectItems[bufID] = obj
	vm.jsPropertyItems[bufID] = make(map[string]jsPropertyDescriptor, 10)

	return Value{Type: VTJSObject, Num: bufID}
}

// jsCallBufferMethod handles static and instance Buffer methods.
// Returns (value, handled bool).
func (vm *VM) jsCallBufferMethod(methodName string, args []Value) (Value, bool) {
	lower := strings.ToLower(methodName)

	switch lower {
	case "from":
		// Buffer.from(data, [encoding]) - create buffer from string or array
		if len(args) == 0 {
			vm.jsThrowTypeError("Buffer.from() requires at least one argument")
			return Value{Type: VTJSUndefined}, true
		}

		encoding := "utf8"
		if len(args) > 1 && args[1].Type != VTJSUndefined {
			encoding = strings.ToLower(vm.valueToString(args[1]))
		}

		var bufData []byte
		source := args[0]

		switch source.Type {
		case VTString:
			// Convert string to bytes with specified encoding
			bufData = vm.jsStringToBuffer(source.Str, encoding)
			if bufData == nil {
				vm.jsThrowTypeError(fmt.Sprintf("Invalid encoding: %s", encoding))
				return Value{Type: VTJSUndefined}, true
			}

		case VTArray:
			// Array-like: convert each element to byte
			if source.Arr != nil {
				bufData = make([]byte, len(source.Arr.Values))
				for i, v := range source.Arr.Values {
					bufData[i] = byte(int(vm.jsToNumber(v).Flt) & 0xFF)
				}
			}

		case VTInteger:
			// Uint8Array-like: allocate new buffer of size
			size := int(source.Num)
			if size < 0 {
				vm.jsThrowRangeError("Buffer size must be non-negative")
				return Value{Type: VTJSUndefined}, true
			}
			bufData = make([]byte, size)

		default:
			vm.jsThrowTypeError("Buffer.from() requires a string, array, or size")
			return Value{Type: VTJSUndefined}, true
		}

		return vm.jsCreateBufferInstance(bufData), true

	case "alloc":
		// Buffer.alloc(size, [fill], [encoding]) - allocate new buffer
		if len(args) == 0 {
			vm.jsThrowTypeError("Buffer.alloc() requires a size argument")
			return Value{Type: VTJSUndefined}, true
		}

		size := int(vm.jsToNumber(args[0]).Flt)
		if size < 0 {
			vm.jsThrowRangeError("Buffer size must be non-negative")
			return Value{Type: VTJSUndefined}, true
		}

		// Check memory limits
		if !vm.jsEnsureStringSize(size) {
			vm.jsThrowRangeError("Buffer allocation exceeds memory limit")
			return Value{Type: VTJSUndefined}, true
		}

		bufData := make([]byte, size)

		// Fill if specified
		if len(args) > 1 && args[1].Type != VTJSUndefined {
			fill := args[1]
			fillByte := byte(0)

			if fill.Type == VTString && len(fill.Str) > 0 {
				fillByte = fill.Str[0]
			} else if fill.Type == VTInteger || fill.Type == VTDouble {
				fillByte = byte(int(vm.jsToNumber(fill).Flt) & 0xFF)
			}

			for i := 0; i < size; i++ {
				bufData[i] = fillByte
			}
		}

		return vm.jsCreateBufferInstance(bufData), true

	case "isbuffer":
		// Buffer.isBuffer(obj) - check if obj is a Buffer
		if len(args) == 0 {
			return Value{Type: VTBool, Num: 0}, true
		}

		obj := args[0]
		if obj.Type == VTJSObject {
			if objType := vm.jsObjectStringProperty(obj, "__js_type"); objType == "Buffer" {
				return Value{Type: VTBool, Num: 1}, true
			}
		}

		return Value{Type: VTBool, Num: 0}, true
	}

	return Value{Type: VTJSUndefined}, false
}

// jsCallBufferInstanceMethod handles Buffer instance methods like toString(), read operations.
// Returns (value, handled bool).
func (vm *VM) jsCallBufferInstanceMethod(bufObj Value, methodName string, args []Value) (Value, bool) {
	lower := strings.ToLower(methodName)

	if bufObj.Type != VTJSObject {
		return Value{Type: VTJSUndefined}, false
	}

	// Get buffer data
	bufferID := bufObj.Num
	bufItem, exists := vm.jsBufferItems[bufferID]
	if !exists {
		return Value{Type: VTJSUndefined}, false
	}

	switch lower {
	case "tostring":
		// buffer.toString([encoding], [start], [end])
		encoding := "utf8"
		if len(args) > 0 && args[0].Type != VTJSUndefined {
			encoding = strings.ToLower(vm.valueToString(args[0]))
		}

		start := 0
		if len(args) > 1 && args[1].Type != VTJSUndefined {
			start = int(vm.jsToNumber(args[1]).Flt)
			if start < 0 {
				start = 0
			}
			if start > len(bufItem.data) {
				start = len(bufItem.data)
			}
		}

		end := len(bufItem.data)
		if len(args) > 2 && args[2].Type != VTJSUndefined {
			end = int(vm.jsToNumber(args[2]).Flt)
			if end < 0 {
				end = 0
			}
			if end > len(bufItem.data) {
				end = len(bufItem.data)
			}
		}

		if start >= end {
			return NewString(""), true
		}

		slice := bufItem.data[start:end]
		result := ""

		switch encoding {
		case "utf8":
			result = string(slice)
		case "hex":
			result = hex.EncodeToString(slice)
		case "base64":
			result = base64.StdEncoding.EncodeToString(slice)
		default:
			vm.jsThrowTypeError(fmt.Sprintf("Invalid encoding: %s", encoding))
			return Value{Type: VTJSUndefined}, true
		}

		return NewString(result), true

	case "length":
		// buffer.length (property access, but included for completeness)
		return NewInteger(int64(len(bufItem.data))), true
	}

	return Value{Type: VTJSUndefined}, false
}

// jsStringToBuffer converts a string to a byte slice using the specified encoding.
// Returns nil if encoding is unsupported.
func (vm *VM) jsStringToBuffer(str string, encoding string) []byte {
	switch strings.ToLower(encoding) {
	case "utf8", "utf-8":
		return []byte(str)

	case "hex":
		data, err := hex.DecodeString(str)
		if err != nil {
			return nil
		}
		return data

	case "base64":
		data, err := base64.StdEncoding.DecodeString(str)
		if err != nil {
			return nil
		}
		return data

	default:
		return nil
	}
}
