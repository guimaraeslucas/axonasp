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
	"encoding/binary"
	"math"
	"math/big"
	"strconv"
	"strings"
)

// jsNumberToUint32 applies ECMAScript ToUint32 semantics to a float64 value.
func jsNumberToUint32(n float64) uint32 {
	if math.IsNaN(n) || math.IsInf(n, 0) || n == 0 {
		return 0
	}
	truncated := math.Trunc(n)
	mod := math.Mod(truncated, 4294967296)
	if mod < 0 {
		mod += 4294967296
	}
	return uint32(mod)
}

// jsNumberToInt32 applies ECMAScript ToInt32 semantics to a float64 value.
func jsNumberToInt32(n float64) int32 {
	u := jsNumberToUint32(n)
	if u >= 0x80000000 {
		return int32(int64(u) - 0x100000000)
	}
	return int32(u)
}

// ---------------------------------------------------------------------------
// Well-Known Symbol IDs
// Negative IDs are reserved for well-known symbols so they never collide
// with user-created symbols (which start at 1 and increment positively).
// ---------------------------------------------------------------------------

const (
	jsWellKnownSymbolIterator           int64 = -1
	jsWellKnownSymbolToStringTag        int64 = -2
	jsWellKnownSymbolSpecies            int64 = -3
	jsWellKnownSymbolHasInstance        int64 = -4
	jsWellKnownSymbolToPrimitive        int64 = -5
	jsWellKnownSymbolDispose            int64 = -6
	jsWellKnownSymbolAsyncDispose       int64 = -7
	jsWellKnownSymbolUnscopables        int64 = -8
	jsWellKnownSymbolMatchAll           int64 = -9
	jsWellKnownSymbolIsConcatSpreadable int64 = -10
)

// jsWellKnownSymbolValue returns a pre-constructed Value for a well-known symbol ID.
func jsWellKnownSymbolValue(id int64, name string) Value {
	return Value{Type: VTSymbol, Num: id, Str: name}
}

// jsIsTypedArrayType returns true if the given __js_type string is one of the
// standard ECMAScript typed array or buffer view types.
func jsIsTypedArrayType(t string) bool {
	switch t {
	case "Int8Array", "Uint8Array", "Uint8ClampedArray",
		"Int16Array", "Uint16Array",
		"Int32Array", "Uint32Array",
		"Float32Array", "Float64Array",
		"BigInt64Array", "BigUint64Array",
		"DataView":
		return true
	}
	return false
}

// jsTypedArrayElementSize returns the byte size of each element for a given typed array type.
func jsTypedArrayElementSize(typeName string) int {
	switch typeName {
	case "Int8Array", "Uint8Array", "Uint8ClampedArray":
		return 1
	case "Int16Array", "Uint16Array":
		return 2
	case "Int32Array", "Uint32Array", "Float32Array":
		return 4
	case "Float64Array", "BigInt64Array", "BigUint64Array":
		return 8
	default:
		return 1
	}
}

// jsGetArrayBufferBytes retrieves the underlying byte slice for an ArrayBuffer or SharedArrayBuffer object.
// Returns nil if the value is not a valid buffer.
func (vm *VM) jsGetArrayBufferBytes(bufObj Value) []byte {
	if bufObj.Type != VTJSObject {
		return nil
	}
	if buf, ok := vm.jsArrayBuffers[bufObj.Num]; ok {
		return buf
	}
	if buf, ok := vm.jsSharedArrayBuffers[bufObj.Num]; ok {
		return buf
	}
	return nil
}

// jsGetTypedArrayInfo extracts the buffer, byteOffset, byteLength, and element size
// from a typed array or DataView object. Returns false if the object is not valid.
func (vm *VM) jsGetTypedArrayInfo(obj Value) (buf []byte, byteOffset, byteLength, elemSize int, ok bool) {
	if obj.Type != VTJSObject {
		return nil, 0, 0, 0, false
	}
	items, exists := vm.jsObjectItems[obj.Num]
	if !exists {
		return nil, 0, 0, 0, false
	}
	typeName, _ := items["__js_type"]
	if !jsIsTypedArrayType(typeName.Str) {
		return nil, 0, 0, 0, false
	}
	bufIDVal, hasBufID := items["__js_buffer_id"]
	if !hasBufID {
		return nil, 0, 0, 0, false
	}
	bufID := bufIDVal.Num
	bufBytes, hasBuf := vm.jsArrayBuffers[bufID]
	if !hasBuf {
		bufBytes, hasBuf = vm.jsSharedArrayBuffers[bufID]
	}
	if !hasBuf {
		return nil, 0, 0, 0, false
	}
	offsetVal := items["__js_byte_offset"]
	offset := int(offsetVal.Num)
	lenVal := items["__js_byte_length"]
	length := int(lenVal.Num)
	if length < 0 {
		length = len(bufBytes) - offset
	}
	eSize := jsTypedArrayElementSize(typeName.Str)
	return bufBytes, offset, length, eSize, true
}

// ---------------------------------------------------------------------------
// ArrayBuffer constructor: new ArrayBuffer(byteLength)
// ---------------------------------------------------------------------------

// jsNewArrayBuffer creates a new ArrayBuffer JS object backed by a Go []byte of the given size.
func (vm *VM) jsNewArrayBuffer(byteLength int) Value {
	if byteLength < 0 {
		vm.jsThrowRangeError("Invalid ArrayBuffer length")
		return Value{Type: VTJSUndefined}
	}
	buf := make([]byte, byteLength)
	objID := vm.allocJSID()
	obj := make(map[string]Value, 4)
	obj["__js_type"] = NewString("ArrayBuffer")
	obj["__js_ctor"] = NewString("ArrayBuffer")
	vm.jsObjectItems[objID] = obj
	vm.jsPropertyItems[objID] = make(map[string]jsPropertyDescriptor, 4)
	vm.jsArrayBuffers[objID] = buf
	return Value{Type: VTJSObject, Num: objID}
}

// ---------------------------------------------------------------------------
// SharedArrayBuffer constructor: new SharedArrayBuffer(byteLength)
// ---------------------------------------------------------------------------

// jsNewSharedArrayBuffer creates a new SharedArrayBuffer JS object backed by a Go []byte.
func (vm *VM) jsNewSharedArrayBuffer(byteLength int) Value {
	if byteLength < 0 {
		vm.jsThrowRangeError("Invalid SharedArrayBuffer length")
		return Value{Type: VTJSUndefined}
	}
	buf := make([]byte, byteLength)
	objID := vm.allocJSID()
	obj := make(map[string]Value, 4)
	obj["__js_type"] = NewString("SharedArrayBuffer")
	obj["__js_ctor"] = NewString("SharedArrayBuffer")
	vm.jsObjectItems[objID] = obj
	vm.jsPropertyItems[objID] = make(map[string]jsPropertyDescriptor, 4)
	vm.jsSharedArrayBuffers[objID] = buf
	return Value{Type: VTJSObject, Num: objID}
}

// ---------------------------------------------------------------------------
// TypedArray constructor: new Uint8Array(length | ArrayBuffer [, offset [, length]])
// ---------------------------------------------------------------------------

// jsNewTypedArray creates a new typed array view over an ArrayBuffer.
// typeName is one of "Int8Array", "Uint8Array", etc.
func (vm *VM) jsNewTypedArray(typeName string, args []Value) Value {
	elemSize := jsTypedArrayElementSize(typeName)
	var bufID int64
	var byteOffset, byteLength int

	if len(args) == 0 {
		// new Uint8Array() → empty, length 0
		bufObj := vm.jsNewArrayBuffer(0)
		bufID = bufObj.Num
		byteOffset = 0
		byteLength = 0
	} else if args[0].Type == VTJSObject {
		items, ok := vm.jsObjectItems[args[0].Num]
		if !ok {
			vm.jsThrowTypeError("TypedArray argument must be an ArrayBuffer or length")
			return Value{Type: VTJSUndefined}
		}
		t, _ := items["__js_type"]
		if t.Str == "ArrayBuffer" || t.Str == "SharedArrayBuffer" {
			// View over existing buffer
			bufID = args[0].Num
			backing := vm.jsGetArrayBufferBytes(args[0])
			if backing == nil {
				vm.jsThrowTypeError("Invalid buffer")
				return Value{Type: VTJSUndefined}
			}
			byteOffset = 0
			if len(args) > 1 {
				byteOffset = int(vm.jsToNumber(args[1]).Flt)
			}
			byteLength = len(backing) - byteOffset
			if len(args) > 2 {
				byteLength = int(vm.jsToNumber(args[2]).Flt) * elemSize
			}
		} else {
			// Treat as array-like source: copy values
			length, hasLen, deferred := vm.jsArrayLikeLength(args[0])
			if deferred {
				return Value{Type: VTJSUndefined}
			}
			if !hasLen {
				length = 0
			}
			bufObj := vm.jsNewArrayBuffer(length * elemSize)
			bufID = bufObj.Num
			byteOffset = 0
			byteLength = length * elemSize
			// Copy elements
			backing := vm.jsArrayBuffers[bufID]
			for i := 0; i < length; i++ {
				v, _ := vm.jsArrayLikeGetIndex(args[0], i)
				vm.jsWriteTypedArrayElement(typeName, elemSize, backing, i, v)
			}
		}
	} else if args[0].Type == VTArray && args[0].Arr != nil {
		// new Uint8Array([1,2,3]) - from JS array
		src := args[0].Arr.Values
		bufObj := vm.jsNewArrayBuffer(len(src) * elemSize)
		bufID = bufObj.Num
		byteOffset = 0
		byteLength = len(src) * elemSize
		backing := vm.jsArrayBuffers[bufID]
		for i, v := range src {
			vm.jsWriteTypedArrayElement(typeName, elemSize, backing, i, v)
		}
	} else {
		// new Uint8Array(length)
		length := int(vm.jsToNumber(args[0]).Flt)
		if length < 0 {
			vm.jsThrowRangeError("Invalid typed array length")
			return Value{Type: VTJSUndefined}
		}
		bufObj := vm.jsNewArrayBuffer(length * elemSize)
		bufID = bufObj.Num
		byteOffset = 0
		byteLength = length * elemSize
	}

	objID := vm.allocJSID()
	obj := make(map[string]Value, 6)
	obj["__js_type"] = NewString(typeName)
	obj["__js_ctor"] = NewString(typeName)
	obj["__js_buffer_id"] = NewInteger(bufID)
	obj["__js_byte_offset"] = NewInteger(int64(byteOffset))
	obj["__js_byte_length"] = NewInteger(int64(byteLength))
	if proto := vm.jsGetIntrinsicPrototype(typeName); proto.Type == VTJSObject {
		obj["__js_proto"] = proto
	}
	vm.jsObjectItems[objID] = obj
	vm.jsPropertyItems[objID] = make(map[string]jsPropertyDescriptor, 4)
	return Value{Type: VTJSObject, Num: objID}
}

// ---------------------------------------------------------------------------
// DataView constructor: new DataView(arrayBuffer [, byteOffset [, byteLength]])
// ---------------------------------------------------------------------------

// jsNewDataView creates a new DataView over an existing ArrayBuffer.
func (vm *VM) jsNewDataView(args []Value) Value {
	if len(args) == 0 || args[0].Type != VTJSObject {
		vm.jsThrowTypeError("DataView requires a buffer argument")
		return Value{Type: VTJSUndefined}
	}
	bufID := args[0].Num
	backing := vm.jsGetArrayBufferBytes(args[0])
	if backing == nil {
		vm.jsThrowTypeError("DataView: argument is not a buffer")
		return Value{Type: VTJSUndefined}
	}
	byteOffset := 0
	byteLength := len(backing)
	if len(args) > 1 {
		byteOffset = int(vm.jsToNumber(args[1]).Flt)
	}
	if len(args) > 2 {
		byteLength = int(vm.jsToNumber(args[2]).Flt)
	} else {
		byteLength = len(backing) - byteOffset
	}
	if byteOffset < 0 || byteOffset > len(backing) {
		vm.jsThrowRangeError("DataView: byteOffset out of range")
		return Value{Type: VTJSUndefined}
	}
	if byteLength < 0 || byteOffset+byteLength > len(backing) {
		vm.jsThrowRangeError("DataView: byteLength out of range")
		return Value{Type: VTJSUndefined}
	}
	objID := vm.allocJSID()
	obj := make(map[string]Value, 6)
	obj["__js_type"] = NewString("DataView")
	obj["__js_ctor"] = NewString("DataView")
	obj["__js_buffer_id"] = NewInteger(bufID)
	obj["__js_byte_offset"] = NewInteger(int64(byteOffset))
	obj["__js_byte_length"] = NewInteger(int64(byteLength))
	vm.jsObjectItems[objID] = obj
	vm.jsPropertyItems[objID] = make(map[string]jsPropertyDescriptor, 4)
	return Value{Type: VTJSObject, Num: objID}
}

// ---------------------------------------------------------------------------
// Typed Array Element Read/Write (safe: bounds panics are recovered by VM)
// ---------------------------------------------------------------------------

// jsReadTypedArrayElement reads element at logical index i from a typed array.
// Out-of-bounds access panics naturally; the VM's defer/recover converts it to RangeError.
func jsReadTypedArrayElement(typeName string, elemSize int, buf []byte, byteOffset, i int) Value {
	base := byteOffset + i*elemSize
	switch typeName {
	case "Int8Array":
		return NewInteger(int64(int8(buf[base])))
	case "Uint8Array", "Uint8ClampedArray":
		return NewInteger(int64(buf[base]))
	case "Int16Array":
		v := binary.LittleEndian.Uint16(buf[base : base+2])
		return NewInteger(int64(int16(v)))
	case "Uint16Array":
		v := binary.LittleEndian.Uint16(buf[base : base+2])
		return NewInteger(int64(v))
	case "Int32Array":
		v := binary.LittleEndian.Uint32(buf[base : base+4])
		return NewInteger(int64(int32(v)))
	case "Uint32Array":
		v := binary.LittleEndian.Uint32(buf[base : base+4])
		return NewInteger(int64(v))
	case "Float32Array":
		bits := binary.LittleEndian.Uint32(buf[base : base+4])
		return NewDouble(float64(math.Float32frombits(bits)))
	case "Float64Array":
		bits := binary.LittleEndian.Uint64(buf[base : base+8])
		return NewDouble(math.Float64frombits(bits))
	case "BigInt64Array":
		bits := binary.LittleEndian.Uint64(buf[base : base+8])
		return NewBigInt(big.NewInt(int64(bits)))
	case "BigUint64Array":
		bits := binary.LittleEndian.Uint64(buf[base : base+8])
		return NewBigInt(new(big.Int).SetUint64(bits))
	}
	return Value{Type: VTJSUndefined}
}

// jsWriteTypedArrayElement writes a numeric value or BigInt to element at logical index i.
// Out-of-bounds access panics naturally; the VM's defer/recover converts it to RangeError.
func (vm *VM) jsWriteTypedArrayElement(typeName string, elemSize int, buf []byte, i int, val Value) {
	base := i * elemSize
	switch typeName {
	case "Int8Array":
		buf[base] = byte(int8(jsNumberToInt32(vm.jsToNumber(val).Flt)))
	case "Uint8Array":
		buf[base] = byte(uint8(jsNumberToUint32(vm.jsToNumber(val).Flt)))
	case "Uint8ClampedArray":
		fval := vm.jsToNumber(val).Flt
		v := int32(fval)
		if v < 0 {
			v = 0
		} else if v > 255 {
			v = 255
		}
		buf[base] = byte(v)
	case "Int16Array":
		binary.LittleEndian.PutUint16(buf[base:base+2], uint16(int16(jsNumberToInt32(vm.jsToNumber(val).Flt))))
	case "Uint16Array":
		binary.LittleEndian.PutUint16(buf[base:base+2], uint16(jsNumberToUint32(vm.jsToNumber(val).Flt)))
	case "Int32Array":
		binary.LittleEndian.PutUint32(buf[base:base+4], uint32(jsNumberToInt32(vm.jsToNumber(val).Flt)))
	case "Uint32Array":
		binary.LittleEndian.PutUint32(buf[base:base+4], jsNumberToUint32(vm.jsToNumber(val).Flt))
	case "Float32Array":
		binary.LittleEndian.PutUint32(buf[base:base+4], math.Float32bits(float32(vm.jsToNumber(val).Flt)))
	case "Float64Array":
		binary.LittleEndian.PutUint64(buf[base:base+8], math.Float64bits(vm.jsToNumber(val).Flt))
	case "BigInt64Array":
		var bits uint64
		if val.Type == VTJSBigInt {
			bits = uint64(val.Big.Int64())
		} else {
			// Coerce to BigInt if possible or throw
			vm.jsThrowTypeError("Cannot convert to BigInt")
			return
		}
		binary.LittleEndian.PutUint64(buf[base:base+8], bits)
	case "BigUint64Array":
		var bits uint64
		if val.Type == VTJSBigInt {
			bits = val.Big.Uint64()
		} else {
			vm.jsThrowTypeError("Cannot convert to BigInt")
			return
		}
		binary.LittleEndian.PutUint64(buf[base:base+8], bits)
	}
}

// ---------------------------------------------------------------------------
// Typed Array: index get/set
// ---------------------------------------------------------------------------

// jsTypedArrayIndexGet retrieves the value at numeric index i from a typed array object.
// Returns (value, true) if handled, (undefined, false) otherwise.
func (vm *VM) jsTypedArrayIndexGet(obj Value, i int) (Value, bool) {
	buf, byteOffset, byteLength, elemSize, ok := vm.jsGetTypedArrayInfo(obj)
	if !ok {
		return Value{Type: VTJSUndefined}, false
	}
	items := vm.jsObjectItems[obj.Num]
	typeName := items["__js_type"].Str
	if typeName == "DataView" {
		return Value{Type: VTJSUndefined}, false
	}
	length := byteLength / elemSize
	if i < 0 || i >= length {
		return Value{Type: VTJSUndefined}, true
	}
	return jsReadTypedArrayElement(typeName, elemSize, buf, byteOffset, i), true
}

// jsTypedArrayIndexSet writes a value at numeric index i to a typed array object.
// Returns true if handled.
func (vm *VM) jsTypedArrayIndexSet(obj Value, i int, val Value) bool {
	buf, byteOffset, byteLength, elemSize, ok := vm.jsGetTypedArrayInfo(obj)
	if !ok {
		return false
	}
	items := vm.jsObjectItems[obj.Num]
	typeName := items["__js_type"].Str
	if typeName == "DataView" {
		return false
	}
	length := byteLength / elemSize
	if i < 0 || i >= length {
		return true // silently ignore out-of-range like spec says
	}
	vm.jsWriteTypedArrayElement(typeName, elemSize, buf[byteOffset:], i, val)
	return true
}

// ---------------------------------------------------------------------------
// Typed Array / DataView: member get (properties)
// ---------------------------------------------------------------------------

// jsTypedArrayMemberGet resolves a named property on a typed array or DataView.
// Returns (value, true) if handled, (undefined, false) otherwise.
func (vm *VM) jsTypedArrayMemberGet(obj Value, member string) (Value, bool) {
	buf, byteOffset, byteLength, elemSize, ok := vm.jsGetTypedArrayInfo(obj)
	if !ok {
		return Value{Type: VTJSUndefined}, false
	}
	items := vm.jsObjectItems[obj.Num]
	typeName := items["__js_type"].Str

	switch {
	case strings.EqualFold(member, "byteLength"):
		return NewInteger(int64(byteLength)), true
	case strings.EqualFold(member, "byteOffset"):
		return NewInteger(int64(byteOffset)), true
	case strings.EqualFold(member, "buffer"):
		bufIDVal := items["__js_buffer_id"]
		return Value{Type: VTJSObject, Num: bufIDVal.Num}, true
	case strings.EqualFold(member, "length") && typeName != "DataView":
		return NewInteger(int64(byteLength / elemSize)), true
	}

	// Numeric index access via string key (e.g., "0", "1")
	if typeName != "DataView" {
		if idx, ok2 := jsParseArrayIndex(member); ok2 {
			length := byteLength / elemSize
			if idx < 0 || idx >= length {
				return Value{Type: VTJSUndefined}, true
			}
			return jsReadTypedArrayElement(typeName, elemSize, buf, byteOffset, idx), true
		}
	}

	return Value{Type: VTJSUndefined}, false
}

// ---------------------------------------------------------------------------
// Typed Array: member set (indexed via string key)
// ---------------------------------------------------------------------------

// jsTypedArrayMemberSet writes to a named numeric property on a typed array.
// Returns true if handled.
func (vm *VM) jsTypedArrayMemberSet(obj Value, member string, val Value) bool {
	buf, byteOffset, byteLength, elemSize, ok := vm.jsGetTypedArrayInfo(obj)
	if !ok {
		return false
	}
	items := vm.jsObjectItems[obj.Num]
	typeName := items["__js_type"].Str
	if typeName == "DataView" {
		return false
	}
	if idx, ok2 := jsParseArrayIndex(member); ok2 {
		length := byteLength / elemSize
		if idx < 0 || idx >= length {
			return true // silently ignore
		}
		vm.jsWriteTypedArrayElement(typeName, elemSize, buf[byteOffset:], idx, val)
		return true
	}
	return false
}

// jsParseArrayIndex tries to parse a string as a non-negative integer array index.
// Returns (index, true) if valid according to ECMAScript rules (no leading zeros except "0", < 2^32-1).
func jsParseArrayIndex(s string) (int, bool) {
	if s == "" || s[0] < '0' || s[0] > '9' {
		return 0, false
	}
	if len(s) > 1 && s[0] == '0' {
		return 0, false
	}
	n, err := strconv.ParseUint(s, 10, 32)
	if err != nil || n >= 0xffffffff {
		return 0, false
	}
	return int(n), true
}

// ---------------------------------------------------------------------------
// DataView: method call dispatch
// ---------------------------------------------------------------------------

// jsDataViewCallMember handles method calls on a DataView object.
// Returns (result, true) if handled, (undefined, false) otherwise.
func (vm *VM) jsDataViewCallMember(obj Value, member string, args []Value) (Value, bool) {
	buf, byteOffset, byteLength, _, ok := vm.jsGetTypedArrayInfo(obj)
	if !ok {
		return Value{Type: VTJSUndefined}, false
	}
	items := vm.jsObjectItems[obj.Num]
	typeName := items["__js_type"].Str
	if typeName != "DataView" {
		return Value{Type: VTJSUndefined}, false
	}

	// Determine the byte offset argument position within the view.
	argByteOffset := 0
	if len(args) > 0 {
		argByteOffset = int(vm.jsToNumber(args[0]).Flt)
	}
	// Absolute position in the backing buffer.
	absOffset := byteOffset + argByteOffset

	// littleEndian flag (second arg for multi-byte methods, true by default here)
	littleEndian := false
	if len(args) > 1 {
		littleEndian = args[1].Type == VTBool && args[1].Num != 0
	}

	// Validate the computed absolute offset is within the DataView's range.
	checkRange := func(needed int) bool {
		return argByteOffset >= 0 && argByteOffset+needed <= byteLength
	}

	// Use platform byte order helpers.
	endianU16 := func(v uint16) uint16 {
		if littleEndian {
			return v
		}
		return (v>>8)&0xFF | (v&0xFF)<<8
	}
	endianU32 := func(v uint32) uint32 {
		if littleEndian {
			return v
		}
		return (v>>24)&0xFF | ((v>>16)&0xFF)<<8 | ((v>>8)&0xFF)<<16 | (v&0xFF)<<24
	}
	endianU64 := func(v uint64) uint64 {
		if littleEndian {
			return v
		}
		var out uint64
		for i := range 8 {
			out |= ((v >> (i * 8)) & 0xFF) << ((7 - i) * 8)
		}
		return out
	}

	setVal := func(needed int) float64 {
		if len(args) < 2 {
			return 0
		}
		// For setters, value is the 2nd arg, littleEndian is 3rd
		return vm.jsToNumber(args[1]).Flt
	}
	setLE := func() bool {
		if len(args) > 2 {
			return args[2].Type == VTBool && args[2].Num != 0
		}
		return false
	}

	switch {
	// ----- Getters -----
	case strings.EqualFold(member, "getInt8"):
		if !checkRange(1) {
			vm.jsThrowRangeError("DataView.getInt8: offset out of range")
			return Value{Type: VTJSUndefined}, true
		}
		return NewInteger(int64(int8(buf[absOffset]))), true
	case strings.EqualFold(member, "getUint8"):
		if !checkRange(1) {
			vm.jsThrowRangeError("DataView.getUint8: offset out of range")
			return Value{Type: VTJSUndefined}, true
		}
		return NewInteger(int64(buf[absOffset])), true
	case strings.EqualFold(member, "getInt16"):
		if !checkRange(2) {
			vm.jsThrowRangeError("DataView.getInt16: offset out of range")
			return Value{Type: VTJSUndefined}, true
		}
		raw := binary.LittleEndian.Uint16(buf[absOffset : absOffset+2])
		return NewInteger(int64(int16(endianU16(raw)))), true
	case strings.EqualFold(member, "getUint16"):
		if !checkRange(2) {
			vm.jsThrowRangeError("DataView.getUint16: offset out of range")
			return Value{Type: VTJSUndefined}, true
		}
		raw := binary.LittleEndian.Uint16(buf[absOffset : absOffset+2])
		return NewInteger(int64(endianU16(raw))), true
	case strings.EqualFold(member, "getInt32"):
		if !checkRange(4) {
			vm.jsThrowRangeError("DataView.getInt32: offset out of range")
			return Value{Type: VTJSUndefined}, true
		}
		raw := binary.LittleEndian.Uint32(buf[absOffset : absOffset+4])
		return NewInteger(int64(int32(endianU32(raw)))), true
	case strings.EqualFold(member, "getUint32"):
		if !checkRange(4) {
			vm.jsThrowRangeError("DataView.getUint32: offset out of range")
			return Value{Type: VTJSUndefined}, true
		}
		raw := binary.LittleEndian.Uint32(buf[absOffset : absOffset+4])
		return NewInteger(int64(endianU32(raw))), true
	case strings.EqualFold(member, "getFloat32"):
		if !checkRange(4) {
			vm.jsThrowRangeError("DataView.getFloat32: offset out of range")
			return Value{Type: VTJSUndefined}, true
		}
		raw := binary.LittleEndian.Uint32(buf[absOffset : absOffset+4])
		return NewDouble(float64(math.Float32frombits(endianU32(raw)))), true
	case strings.EqualFold(member, "getFloat64"):
		if !checkRange(8) {
			vm.jsThrowRangeError("DataView.getFloat64: offset out of range")
			return Value{Type: VTJSUndefined}, true
		}
		raw := binary.LittleEndian.Uint64(buf[absOffset : absOffset+8])
		return NewDouble(math.Float64frombits(endianU64(raw))), true

	// ----- Setters -----
	case strings.EqualFold(member, "setInt8"):
		if !checkRange(1) {
			vm.jsThrowRangeError("DataView.setInt8: offset out of range")
			return Value{Type: VTJSUndefined}, true
		}
		buf[absOffset] = byte(int8(jsNumberToInt32(setVal(1))))
		return Value{Type: VTJSUndefined}, true
	case strings.EqualFold(member, "setUint8"):
		if !checkRange(1) {
			vm.jsThrowRangeError("DataView.setUint8: offset out of range")
			return Value{Type: VTJSUndefined}, true
		}
		buf[absOffset] = byte(uint8(jsNumberToUint32(setVal(1))))
		return Value{Type: VTJSUndefined}, true
	case strings.EqualFold(member, "setInt16"):
		if !checkRange(2) {
			vm.jsThrowRangeError("DataView.setInt16: offset out of range")
			return Value{Type: VTJSUndefined}, true
		}
		v := uint16(int16(jsNumberToInt32(setVal(2))))
		le := setLE()
		if le {
			binary.LittleEndian.PutUint16(buf[absOffset:absOffset+2], v)
		} else {
			binary.BigEndian.PutUint16(buf[absOffset:absOffset+2], v)
		}
		return Value{Type: VTJSUndefined}, true
	case strings.EqualFold(member, "setUint16"):
		if !checkRange(2) {
			vm.jsThrowRangeError("DataView.setUint16: offset out of range")
			return Value{Type: VTJSUndefined}, true
		}
		v := uint16(jsNumberToUint32(setVal(2)))
		le := setLE()
		if le {
			binary.LittleEndian.PutUint16(buf[absOffset:absOffset+2], v)
		} else {
			binary.BigEndian.PutUint16(buf[absOffset:absOffset+2], v)
		}
		return Value{Type: VTJSUndefined}, true
	case strings.EqualFold(member, "setInt32"):
		if !checkRange(4) {
			vm.jsThrowRangeError("DataView.setInt32: offset out of range")
			return Value{Type: VTJSUndefined}, true
		}
		v := uint32(jsNumberToInt32(setVal(4)))
		le := setLE()
		if le {
			binary.LittleEndian.PutUint32(buf[absOffset:absOffset+4], v)
		} else {
			binary.BigEndian.PutUint32(buf[absOffset:absOffset+4], v)
		}
		return Value{Type: VTJSUndefined}, true
	case strings.EqualFold(member, "setUint32"):
		if !checkRange(4) {
			vm.jsThrowRangeError("DataView.setUint32: offset out of range")
			return Value{Type: VTJSUndefined}, true
		}
		v := jsNumberToUint32(setVal(4))
		le := setLE()
		if le {
			binary.LittleEndian.PutUint32(buf[absOffset:absOffset+4], v)
		} else {
			binary.BigEndian.PutUint32(buf[absOffset:absOffset+4], v)
		}
		return Value{Type: VTJSUndefined}, true
	case strings.EqualFold(member, "setFloat32"):
		if !checkRange(4) {
			vm.jsThrowRangeError("DataView.setFloat32: offset out of range")
			return Value{Type: VTJSUndefined}, true
		}
		v := math.Float32bits(float32(setVal(4)))
		le := setLE()
		if le {
			binary.LittleEndian.PutUint32(buf[absOffset:absOffset+4], v)
		} else {
			binary.BigEndian.PutUint32(buf[absOffset:absOffset+4], v)
		}
		return Value{Type: VTJSUndefined}, true
	case strings.EqualFold(member, "setFloat64"):
		if !checkRange(8) {
			vm.jsThrowRangeError("DataView.setFloat64: offset out of range")
			return Value{Type: VTJSUndefined}, true
		}
		v := math.Float64bits(setVal(8))
		le := setLE()
		if le {
			binary.LittleEndian.PutUint64(buf[absOffset:absOffset+8], v)
		} else {
			binary.BigEndian.PutUint64(buf[absOffset:absOffset+8], v)
		}
		return Value{Type: VTJSUndefined}, true
	}

	// Suppress "unused" lint for the helper closures when no method matches.
	_ = endianU16
	_ = endianU32
	_ = endianU64

	return Value{Type: VTJSUndefined}, false
}

// ---------------------------------------------------------------------------
// Typed Array: set() method — copies values from another typed array or array
// ---------------------------------------------------------------------------

// jsTypedArraySet implements typedArray.set(source [, offset]).
func (vm *VM) jsTypedArraySet(obj Value, args []Value) Value {
	buf, byteOffset, byteLength, elemSize, ok := vm.jsGetTypedArrayInfo(obj)
	if !ok {
		return Value{Type: VTJSUndefined}
	}
	items := vm.jsObjectItems[obj.Num]
	typeName := items["__js_type"].Str
	if typeName == "DataView" {
		return Value{Type: VTJSUndefined}
	}

	targetOffset := 0
	if len(args) > 1 {
		targetOffset = int(vm.jsToNumber(args[1]).Flt)
	}
	if len(args) == 0 {
		return Value{Type: VTJSUndefined}
	}

	src := args[0]
	length := byteLength / elemSize

	// If source is another typed array
	srcBuf, srcByteOffset, srcByteLength, srcElemSize, isSrcTyped := vm.jsGetTypedArrayInfo(src)
	if isSrcTyped {
		srcItems := vm.jsObjectItems[src.Num]
		srcTypeName := srcItems["__js_type"].Str
		if srcTypeName != "DataView" {
			srcLen := srcByteLength / srcElemSize
			if targetOffset+srcLen > length {
				vm.jsThrowRangeError("TypedArray.set: source is too large")
				return Value{Type: VTJSUndefined}
			}
			for i := range srcLen {
				v := jsReadTypedArrayElement(srcTypeName, srcElemSize, srcBuf, srcByteOffset, i)
				vm.jsWriteTypedArrayElement(typeName, elemSize, buf[byteOffset:], targetOffset+i, v)
			}
			return Value{Type: VTJSUndefined}
		}
	}

	// Array-like source
	srcLen, hasLen, deferred := vm.jsArrayLikeLength(src)
	if deferred {
		return Value{Type: VTJSUndefined}
	}
	if !hasLen {
		return Value{Type: VTJSUndefined}
	}
	if targetOffset+srcLen > length {
		vm.jsThrowRangeError("TypedArray.set: source is too large")
		return Value{Type: VTJSUndefined}
	}
	for i := range srcLen {
		v, _ := vm.jsArrayLikeGetIndex(src, i)
		vm.jsWriteTypedArrayElement(typeName, elemSize, buf[byteOffset:], targetOffset+i, v)
	}
	return Value{Type: VTJSUndefined}
}

// ---------------------------------------------------------------------------
// Typed Array: subarray() — returns a new view into the same buffer
// ---------------------------------------------------------------------------

// jsTypedArraySubarray implements typedArray.subarray([begin[, end]]).
func (vm *VM) jsTypedArraySubarray(obj Value, args []Value) Value {
	buf, byteOffset, byteLength, elemSize, ok := vm.jsGetTypedArrayInfo(obj)
	if !ok || buf == nil {
		return Value{Type: VTJSUndefined}
	}
	items := vm.jsObjectItems[obj.Num]
	typeName := items["__js_type"].Str
	bufIDVal := items["__js_buffer_id"]
	length := byteLength / elemSize

	begin := 0
	end := length
	if len(args) > 0 {
		begin = int(vm.jsToNumber(args[0]).Flt)
		if begin < 0 {
			begin = length + begin
		}
		if begin < 0 {
			begin = 0
		}
		if begin > length {
			begin = length
		}
	}
	if len(args) > 1 {
		end = int(vm.jsToNumber(args[1]).Flt)
		if end < 0 {
			end = length + end
		}
		if end < 0 {
			end = 0
		}
		if end > length {
			end = length
		}
	}
	if end < begin {
		end = begin
	}

	newByteOffset := byteOffset + begin*elemSize
	newByteLength := (end - begin) * elemSize

	objID := vm.allocJSID()
	newObj := make(map[string]Value, 6)
	newObj["__js_type"] = NewString(typeName)
	newObj["__js_ctor"] = NewString(typeName)
	newObj["__js_buffer_id"] = NewInteger(bufIDVal.Num)
	newObj["__js_byte_offset"] = NewInteger(int64(newByteOffset))
	newObj["__js_byte_length"] = NewInteger(int64(newByteLength))
	vm.jsObjectItems[objID] = newObj
	vm.jsPropertyItems[objID] = make(map[string]jsPropertyDescriptor, 4)
	return Value{Type: VTJSObject, Num: objID}
}

// ---------------------------------------------------------------------------
// Typed Array: fill() — fills the array with a static value
// ---------------------------------------------------------------------------

// jsTypedArrayFill implements typedArray.fill(value[, start[, end]]).
func (vm *VM) jsTypedArrayFill(obj Value, args []Value) Value {
	buf, byteOffset, byteLength, elemSize, ok := vm.jsGetTypedArrayInfo(obj)
	if !ok {
		return obj
	}
	items := vm.jsObjectItems[obj.Num]
	typeName := items["__js_type"].Str
	if typeName == "DataView" {
		return obj
	}
	length := byteLength / elemSize
	start := 0
	end := length
	if len(args) > 1 {
		start = int(vm.jsToNumber(args[1]).Flt)
		if start < 0 {
			start = length + start
		}
	}
	if len(args) > 2 {
		end = int(vm.jsToNumber(args[2]).Flt)
		if end < 0 {
			end = length + end
		}
	}
	if start < 0 {
		start = 0
	}
	if end > length {
		end = length
	}
	for i := start; i < end; i++ {
		vm.jsWriteTypedArrayElement(typeName, elemSize, buf[byteOffset:], i, args[0])
	}
	return obj
}

// ---------------------------------------------------------------------------
// ArrayBuffer: slice() — creates a new ArrayBuffer from a sub-range
// ---------------------------------------------------------------------------

// jsArrayBufferSlice implements arrayBuffer.slice(begin[, end]).
func (vm *VM) jsArrayBufferSlice(obj Value, args []Value) Value {
	if obj.Type != VTJSObject {
		return Value{Type: VTJSUndefined}
	}
	backing := vm.jsGetArrayBufferBytes(obj)
	if backing == nil {
		return Value{Type: VTJSUndefined}
	}
	objType := vm.jsObjectStringProperty(obj, "__js_type")

	length := len(backing)
	begin := 0
	end := length
	if len(args) > 0 {
		begin = int(vm.jsToNumber(args[0]).Flt)
		if begin < 0 {
			begin = length + begin
		}
		if begin < 0 {
			begin = 0
		}
		if begin > length {
			begin = length
		}
	}
	if len(args) > 1 {
		end = int(vm.jsToNumber(args[1]).Flt)
		if end < 0 {
			end = length + end
		}
		if end < 0 {
			end = 0
		}
		if end > length {
			end = length
		}
	}
	if end < begin {
		end = begin
	}
	newBuf := make([]byte, end-begin)
	copy(newBuf, backing[begin:end])

	if objType == "SharedArrayBuffer" {
		return vm.jsNewSharedArrayBufferWithBacking(newBuf)
	}
	return vm.jsNewArrayBufferWithBacking(newBuf)
}

// jsNewArrayBufferWithBacking creates a new ArrayBuffer JS object using an existing byte slice.
func (vm *VM) jsNewArrayBufferWithBacking(buf []byte) Value {
	objID := vm.allocJSID()
	obj := make(map[string]Value, 4)
	obj["__js_type"] = NewString("ArrayBuffer")
	obj["__js_ctor"] = NewString("ArrayBuffer")
	vm.jsObjectItems[objID] = obj
	vm.jsPropertyItems[objID] = make(map[string]jsPropertyDescriptor, 4)
	vm.jsArrayBuffers[objID] = buf
	return Value{Type: VTJSObject, Num: objID}
}

// jsNewSharedArrayBufferWithBacking creates a new SharedArrayBuffer JS object using an existing byte slice.
func (vm *VM) jsNewSharedArrayBufferWithBacking(buf []byte) Value {
	objID := vm.allocJSID()
	obj := make(map[string]Value, 4)
	obj["__js_type"] = NewString("SharedArrayBuffer")
	obj["__js_ctor"] = NewString("SharedArrayBuffer")
	vm.jsObjectItems[objID] = obj
	vm.jsPropertyItems[objID] = make(map[string]jsPropertyDescriptor, 4)
	vm.jsSharedArrayBuffers[objID] = buf
	return Value{Type: VTJSObject, Num: objID}
}

// ---------------------------------------------------------------------------
// Typed Array: for-of iteration support (collect all element values)
// ---------------------------------------------------------------------------

// jsTypedArrayValues returns all element values of a typed array for iteration.
func (vm *VM) jsTypedArrayValues(obj Value) []Value {
	buf, byteOffset, byteLength, elemSize, ok := vm.jsGetTypedArrayInfo(obj)
	if !ok {
		return nil
	}
	items := vm.jsObjectItems[obj.Num]
	typeName := items["__js_type"].Str
	if typeName == "DataView" {
		return nil
	}
	length := byteLength / elemSize
	out := make([]Value, length)
	for i := range length {
		out[i] = jsReadTypedArrayElement(typeName, elemSize, buf, byteOffset, i)
	}
	return out
}

// ---------------------------------------------------------------------------
// Symbol.for / Symbol.keyFor helpers
// ---------------------------------------------------------------------------

// jsSymbolFor implements Symbol.for(key): returns the same symbol for the same key string.
func (vm *VM) jsSymbolFor(key string) Value {
	if sym, ok := vm.jsSymbolGlobalRegistry[key]; ok {
		return sym
	}
	sym := Value{Type: VTSymbol, Num: vm.jsAllocSymbolID(), Str: key}
	vm.jsSymbolGlobalRegistry[key] = sym
	vm.jsRegisteredSymbolIDs[sym.Num] = struct{}{}
	return sym
}

// jsSymbolKeyFor implements Symbol.keyFor(sym): returns the key for a globally registered symbol.
func (vm *VM) jsSymbolKeyFor(sym Value) Value {
	if sym.Type != VTSymbol {
		vm.jsThrowTypeError("Symbol.keyFor: argument is not a Symbol")
		return Value{Type: VTJSUndefined}
	}
	// Well-known symbols are not in the global registry
	for key, s := range vm.jsSymbolGlobalRegistry {
		if s.Num == sym.Num {
			return NewString(key)
		}
	}
	return Value{Type: VTJSUndefined}
}

// ---------------------------------------------------------------------------
// RangeError throw helper (for out-of-range buffer operations)
// ---------------------------------------------------------------------------

// jsThrowRangeError throws a JScript RangeError catchable by a JS try/catch block.
func (vm *VM) jsThrowRangeError(msg string) {
	if len(vm.jsTryStack) == 0 {
		vm.raise(5, msg) // InvalidProcedureCallOrArgument is a reasonable fallback
		return
	}
	target := vm.jsTryStack[len(vm.jsTryStack)-1]
	vm.jsTryStack = vm.jsTryStack[:len(vm.jsTryStack)-1]
	vm.jsErrStack = append(vm.jsErrStack, NewString("RangeError: "+msg))
	vm.ip = target
}
