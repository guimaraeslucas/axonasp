/*
 * AxonASP Server
 * Copyright (C) 2026 G3pix Ltda. All rights reserved.
 *
 * Developed by Lucas Guimaraes - G3pix Ltda
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
	_ "embed"
	"encoding/base64"
	"encoding/hex"
	"strings"
)

//go:embed node_js/stream.js
var jsStreamPolyfillSource string

const jsStreamModuleKey = "__builtin__:stream"

type jsNodeStreamHookResource struct {
	readableChunks [][]byte
	readPos        int
	writable       []byte
	ended          bool
}

// jsGetOrCreateStreamModule returns the cached stream module object.
func (vm *VM) jsGetOrCreateStreamModule() Value {
	if env, ok := vm.jsModuleInstances[jsStreamModuleKey]; ok && env != nil {
		if moduleVal, ok := env.bindings["module"]; ok {
			return moduleVal
		}
	}

	moduleVal := vm.jsRunNodeStreamPolyfill()
	if moduleVal.Type == VTJSUndefined {
		return Value{Type: VTJSUndefined}
	}

	moduleEnv := &jsEnvFrame{
		parentID: vm.jsRootEnvID,
		bindings: map[string]Value{
			"module": moduleVal,
		},
	}
	vm.jsModuleInstances[jsStreamModuleKey] = moduleEnv

	return moduleVal
}

// jsRunNodeStreamPolyfill executes the embedded stream.js polyfill and reads the
// Stream module object from the root scope.
func (vm *VM) jsRunNodeStreamPolyfill() Value {
	source := jsStreamPolyfillSource
	if source == "" {
		return Value{Type: VTJSUndefined}
	}

	compiler := NewASPCompiler("")
	compiler.sourceName = "__builtin__:stream"
	vm.jsPrepareDynamicCompilerIC(compiler)
	compiler.compileJScriptBlock(source)
	compiler.emit(OpHalt)

	if len(compiler.bytecode) == 0 {
		return Value{Type: VTJSUndefined}
	}

	vm.jsExtendICStateFromCompiler(compiler)

	startIP := vm.appendExecuteProgram(compiler.GlobalsCount(), compiler.constants, compiler.bytecode)
	if startIP < 0 || startIP >= len(vm.bytecode) {
		return Value{Type: VTJSUndefined}
	}

	// Use cloneForExecuteGlobal for the same reason as the events polyfill:
	// the polyfill is a standalone script and must not inherit the caller's
	// lexical block scopes.
	child := vm.cloneForExecuteGlobal(startIP)
	child.sourceName = "__builtin__:stream"
	child.jsActiveEnvID = child.jsRootEnvID
	if err := child.Run(); err != nil {
		vm.syncExecuteGlobalState(child)
		return Value{Type: VTJSUndefined}
	}

	vm.syncExecuteGlobalState(child)
	return vm.jsNodeGetRootBinding("Stream")
}

// jsCreateNodeStreamHooksObject allocates the internal stream hook bridge object.
func (vm *VM) jsCreateNodeStreamHooksObject() Value {
	objID := vm.allocJSID()
	obj := make(map[string]Value, 2)
	obj["__js_type"] = NewString("__axon_stream")
	vm.jsObjectItems[objID] = obj
	vm.jsPropertyItems[objID] = make(map[string]jsPropertyDescriptor, 4)
	return Value{Type: VTJSObject, Num: objID}
}

func (vm *VM) jsNodeStreamMakeHandle(resourceID int64) Value {
	handleID := vm.allocJSID()
	handleObj := make(map[string]Value, 4)
	handleObj["__js_type"] = NewString("__axon_stream_handle")
	handleObj["__axon_stream_id"] = NewInteger(resourceID)
	vm.jsObjectItems[handleID] = handleObj
	vm.jsPropertyItems[handleID] = make(map[string]jsPropertyDescriptor, 4)
	return Value{Type: VTJSObject, Num: handleID}
}

func (vm *VM) jsNodeStreamGetResource(handle Value) (*jsNodeStreamHookResource, bool) {
	if handle.Type == VTInteger || handle.Type == VTDouble {
		resourceID := int64(vm.jsToNumber(handle).Flt)
		if resource, ok := vm.jsStreamHookItems[resourceID]; ok && resource != nil {
			return resource, true
		}
		return nil, false
	}
	if handle.Type != VTJSObject && handle.Type != VTJSFunction {
		return nil, false
	}
	idVal, deferred := vm.jsMemberGet(handle, "__axon_stream_id")
	if deferred {
		return nil, false
	}
	if idVal.Type != VTInteger && idVal.Type != VTDouble {
		return nil, false
	}
	resourceID := int64(vm.jsToNumber(idVal).Flt)
	if resource, ok := vm.jsStreamHookItems[resourceID]; ok && resource != nil {
		return resource, true
	}
	return nil, false
}

func (vm *VM) jsNodeStreamCollectSourceChunks(source Value, encoding string) [][]byte {
	if source.Type == VTJSUndefined || source.Type == VTNull {
		return nil
	}
	length, isArrayLike, deferred := vm.jsArrayLikeLength(source)
	if isArrayLike && !deferred {
		chunks := make([][]byte, 0, length)
		for i := range length {
			item, ok := vm.jsArrayLikeGetIndex(source, i)
			if !ok {
				continue
			}
			bytesVal, bytesOK := vm.jsNodeValueBytes(item, encoding)
			if !bytesOK {
				bytesVal = []byte(vm.valueToString(item))
			}
			chunk := make([]byte, len(bytesVal))
			copy(chunk, bytesVal)
			chunks = append(chunks, chunk)
		}
		return chunks
	}
	bytesVal, ok := vm.jsNodeValueBytes(source, encoding)
	if !ok {
		bytesVal = []byte(vm.valueToString(source))
	}
	chunk := make([]byte, len(bytesVal))
	copy(chunk, bytesVal)
	return [][]byte{chunk}
}

func (vm *VM) jsNodeStreamEncodeBytes(data []byte, encoding string) Value {
	enc := strings.ToLower(strings.TrimSpace(encoding))
	if enc == "" {
		return vm.jsCreateBufferInstance(data)
	}
	switch enc {
	case "utf8", "utf-8":
		return NewString(string(data))
	case "hex":
		return NewString(hex.EncodeToString(data))
	case "base64":
		return NewString(base64.StdEncoding.EncodeToString(data))
	default:
		return NewString(string(data))
	}
}

// jsCallNodeStreamHookMethod handles stream native hook invocations.
func (vm *VM) jsCallNodeStreamHookMethod(methodName string, args []Value) (Value, bool) {
	switch strings.ToLower(methodName) {
	case "createreadable":
		source := jsArgOrUndefined(args, 0)
		encoding := vm.valueToString(jsArgOrUndefined(args, 1))
		sourceOptions := source
		if sourceOptions.Type == VTJSObject || sourceOptions.Type == VTJSFunction {
			if sourceOpt, deferred := vm.jsMemberGet(sourceOptions, "source"); !deferred && sourceOpt.Type != VTJSUndefined {
				source = sourceOpt
			}
			if encOpt, deferred := vm.jsMemberGet(sourceOptions, "encoding"); !deferred && encOpt.Type != VTJSUndefined && encoding == "" {
				encoding = vm.valueToString(encOpt)
			}
		}
		resourceID := vm.allocJSID()
		vm.jsStreamHookItems[resourceID] = &jsNodeStreamHookResource{
			readableChunks: vm.jsNodeStreamCollectSourceChunks(source, encoding),
			readPos:        0,
			writable:       make([]byte, 0, 128),
			ended:          true,
		}
		return vm.jsNodeStreamMakeHandle(resourceID), true
	case "createwritable":
		resourceID := vm.allocJSID()
		vm.jsStreamHookItems[resourceID] = &jsNodeStreamHookResource{
			readableChunks: nil,
			readPos:        0,
			writable:       make([]byte, 0, 128),
			ended:          false,
		}
		return vm.jsNodeStreamMakeHandle(resourceID), true
	case "createduplex":
		source := jsArgOrUndefined(args, 0)
		encoding := vm.valueToString(jsArgOrUndefined(args, 1))
		sourceOptions := source
		if sourceOptions.Type == VTJSObject || sourceOptions.Type == VTJSFunction {
			if sourceOpt, deferred := vm.jsMemberGet(sourceOptions, "source"); !deferred && sourceOpt.Type != VTJSUndefined {
				source = sourceOpt
			}
			if encOpt, deferred := vm.jsMemberGet(sourceOptions, "encoding"); !deferred && encOpt.Type != VTJSUndefined && encoding == "" {
				encoding = vm.valueToString(encOpt)
			}
		}
		resourceID := vm.allocJSID()
		vm.jsStreamHookItems[resourceID] = &jsNodeStreamHookResource{
			readableChunks: vm.jsNodeStreamCollectSourceChunks(source, encoding),
			readPos:        0,
			writable:       make([]byte, 0, 128),
			ended:          false,
		}
		return vm.jsNodeStreamMakeHandle(resourceID), true
	case "pull":
		resource, ok := vm.jsNodeStreamGetResource(jsArgOrUndefined(args, 0))
		if !ok {
			return Value{Type: VTJSUndefined}, true
		}
		if resource.readPos >= len(resource.readableChunks) {
			return Value{Type: VTJSUndefined}, true
		}
		chunk := resource.readableChunks[resource.readPos]
		resource.readPos++
		out := make([]byte, len(chunk))
		copy(out, chunk)
		return vm.jsCreateBufferInstance(out), true
	case "write":
		resource, ok := vm.jsNodeStreamGetResource(jsArgOrUndefined(args, 0))
		if !ok {
			return NewInteger(0), true
		}
		chunk := jsArgOrUndefined(args, 1)
		encoding := vm.valueToString(jsArgOrUndefined(args, 2))
		bytesVal, bytesOK := vm.jsNodeValueBytes(chunk, encoding)
		if !bytesOK {
			bytesVal = []byte(vm.valueToString(chunk))
		}
		resource.writable = append(resource.writable, bytesVal...)
		return NewInteger(int64(len(bytesVal))), true
	case "enqueue":
		resource, ok := vm.jsNodeStreamGetResource(jsArgOrUndefined(args, 0))
		if !ok {
			return NewInteger(0), true
		}
		chunk := jsArgOrUndefined(args, 1)
		encoding := vm.valueToString(jsArgOrUndefined(args, 2))
		bytesVal, bytesOK := vm.jsNodeValueBytes(chunk, encoding)
		if !bytesOK {
			bytesVal = []byte(vm.valueToString(chunk))
		}
		queued := make([]byte, len(bytesVal))
		copy(queued, bytesVal)
		resource.readableChunks = append(resource.readableChunks, queued)
		return NewInteger(int64(len(queued))), true
	case "end":
		resource, ok := vm.jsNodeStreamGetResource(jsArgOrUndefined(args, 0))
		if !ok {
			return Value{Type: VTJSUndefined}, true
		}
		if len(args) > 1 && args[1].Type != VTJSUndefined {
			encoding := vm.valueToString(jsArgOrUndefined(args, 2))
			bytesVal, bytesOK := vm.jsNodeValueBytes(args[1], encoding)
			if !bytesOK {
				bytesVal = []byte(vm.valueToString(args[1]))
			}
			resource.writable = append(resource.writable, bytesVal...)
		}
		resource.ended = true
		return Value{Type: VTJSUndefined}, true
	case "eof":
		resource, ok := vm.jsNodeStreamGetResource(jsArgOrUndefined(args, 0))
		if !ok {
			return NewBool(true), true
		}
		return NewBool(resource.readPos >= len(resource.readableChunks)), true
	case "readall":
		resource, ok := vm.jsNodeStreamGetResource(jsArgOrUndefined(args, 0))
		if !ok {
			return Value{Type: VTJSUndefined}, true
		}
		encoding := vm.valueToString(jsArgOrUndefined(args, 1))
		return vm.jsNodeStreamEncodeBytes(resource.writable, encoding), true
	}
	return Value{Type: VTJSUndefined}, false
}
