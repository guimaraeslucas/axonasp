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
	"bytes"
	"crypto/hmac"
	"crypto/md5"
	"crypto/rand"
	"crypto/sha1"
	"crypto/sha256"
	"encoding/base64"
	"encoding/hex"
	"io"
	"net/http"
	"os"
	"path/filepath"
	"strconv"
	"strings"
	"time"
)

const jsAsyncFSReadResultQueueSize = 256
const jsCommonJSExportsCacheKey = "__js_cjs_exports"

type jsAsyncFSReadResult struct {
	promise    Value
	callback   Value
	data       []byte
	encoding   string
	errMsg     string
	asPromise  bool
	callbackOk bool
}

// jsIsPathModuleSpecifier reports whether a require() specifier is a path-like import.
func jsIsPathModuleSpecifier(specifier string) bool {
	if specifier == "" {
		return false
	}
	if filepath.IsAbs(specifier) {
		return true
	}
	if strings.HasPrefix(specifier, "./") || strings.HasPrefix(specifier, "../") {
		return true
	}
	if strings.HasPrefix(specifier, ".\\") || strings.HasPrefix(specifier, "..\\") {
		return true
	}
	if strings.HasPrefix(specifier, "/") || strings.HasPrefix(specifier, "\\") {
		return true
	}
	return false
}

// jsGetCommonJSModuleExports extracts module.exports from one cached CommonJS environment.
func (vm *VM) jsGetCommonJSModuleExports(env *jsEnvFrame) Value {
	if env == nil {
		return Value{Type: VTJSUndefined}
	}
	if exportsVal, ok := env.bindings[jsCommonJSExportsCacheKey]; ok {
		return exportsVal
	}
	moduleVal, ok := env.bindings["module"]
	if !ok {
		return Value{Type: VTJSUndefined}
	}
	exportsVal, deferred := vm.jsMemberGet(moduleVal, "exports")
	if deferred {
		return Value{Type: VTJSUndefined}
	}
	env.bindings[jsCommonJSExportsCacheKey] = exportsVal
	return exportsVal
}

// jsRequireFileModule loads and executes a CommonJS module from disk.
func (vm *VM) jsRequireFileModule(moduleName string) (Value, bool, bool) {
	modulePath, err := vm.jsResolveModulePath(moduleName)
	if err != nil {
		return Value{Type: VTJSUndefined}, false, false
	}

	if env, ok := vm.jsModuleInstances[modulePath]; ok && env != nil {
		return vm.jsGetCommonJSModuleExports(env), true, false
	}

	info, statErr := os.Stat(modulePath)
	if statErr != nil || info.IsDir() {
		return Value{Type: VTJSUndefined}, false, false
	}

	vm.ensureJSRootEnv()
	rootEnvID := vm.jsActiveEnvID

	exportsObjID := vm.allocJSID()
	exportsObj := make(map[string]Value, 8)
	vm.jsObjectItems[exportsObjID] = exportsObj
	vm.jsPropertyItems[exportsObjID] = make(map[string]jsPropertyDescriptor, 8)
	exportsVal := Value{Type: VTJSObject, Num: exportsObjID}

	moduleObjID := vm.allocJSID()
	moduleObj := make(map[string]Value, 4)
	moduleObj["exports"] = exportsVal
	vm.jsObjectItems[moduleObjID] = moduleObj
	vm.jsPropertyItems[moduleObjID] = make(map[string]jsPropertyDescriptor, 4)
	moduleVal := Value{Type: VTJSObject, Num: moduleObjID}

	moduleEnvID := vm.allocJSID()
	moduleEnv := &jsEnvFrame{parentID: rootEnvID, bindings: make(map[string]Value, 12)}
	moduleEnv.bindings["module"] = moduleVal
	moduleEnv.bindings["exports"] = exportsVal
	moduleEnv.bindings["require"] = vm.jsCreateIntrinsicObject("", "require")
	moduleEnv.bindings["__filename"] = NewString(modulePath)
	moduleEnv.bindings["__dirname"] = NewString(filepath.Dir(modulePath))
	vm.jsEnvItems[moduleEnvID] = moduleEnv
	vm.jsModuleInstances[modulePath] = moduleEnv
	vm.jsModuleLoading[modulePath] = struct{}{}
	defer delete(vm.jsModuleLoading, modulePath)

	cache := getExecuteScriptCache()
	program, loadErr := cache.LoadOrCompile(modulePath)
	if loadErr != nil {
		delete(vm.jsModuleInstances, modulePath)
		delete(vm.jsEnvItems, moduleEnvID)
		vm.jsThrowReferenceError("Cannot load module '" + modulePath + "': " + loadErr.Error())
		return Value{Type: VTJSUndefined}, true, true
	}

	startIP := vm.appendExecuteProgram(program.GlobalCount, program.Constants, program.Bytecode)
	child := vm.cloneForExecuteGlobal(startIP)
	child.engineMode = EngineModeJavaScript
	child.sourceName = modulePath
	child.baseSourceName = modulePath
	child.jsActiveEnvID = moduleEnvID
	if runErr := child.Run(); runErr != nil {
		delete(vm.jsModuleInstances, modulePath)
		delete(vm.jsEnvItems, moduleEnvID)
		vm.jsThrowReferenceError("Error executing module '" + modulePath + "': " + runErr.Error())
		return Value{Type: VTJSUndefined}, true, true
	}

	vm.nextDynamicNativeID = child.nextDynamicNativeID
	vm.jsNextSymbolID = child.jsNextSymbolID
	vm.jsRootEnvID = child.jsRootEnvID
	if finalEnv, ok := vm.jsEnvItems[moduleEnvID]; ok && finalEnv != nil {
		vm.jsModuleInstances[modulePath] = finalEnv
		exportsFinal := vm.jsGetCommonJSModuleExports(finalEnv)
		if exportsFinal.Type == VTJSUndefined {
			finalEnv.bindings[jsCommonJSExportsCacheKey] = exportsVal
			return exportsVal, true, false
		}
		return exportsFinal, true, false
	}

	return exportsVal, true, false
}

// jsNodeGetRootBinding reads one binding directly from the JS root environment.
// This bypasses local lexical scopes (including TDZ) and is used by require().
func (vm *VM) jsNodeGetRootBinding(name string) Value {
	vm.ensureJSRootEnv()
	if vm.jsRootEnvID == 0 {
		return Value{Type: VTJSUndefined}
	}
	root := vm.jsEnvItems[vm.jsRootEnvID]
	if root == nil {
		return Value{Type: VTJSUndefined}
	}
	if val, ok := root.bindings[name]; ok {
		return val
	}
	return Value{Type: VTJSUndefined}
}

// jsRequire resolves built-in Node.js-compatible modules exposed by this VM.
func (vm *VM) jsRequire(args []Value) Value {
	if len(args) < 1 {
		vm.jsThrowTypeError("require expects a module name")
		return Value{Type: VTJSUndefined}
	}

	moduleName := strings.TrimSpace(vm.valueToString(args[0]))
	if moduleName == "" {
		vm.jsThrowTypeError("require expects a module name")
		return Value{Type: VTJSUndefined}
	}

	if jsIsPathModuleSpecifier(moduleName) {
		moduleVal, resolved, threw := vm.jsRequireFileModule(moduleName)
		if threw {
			return Value{Type: VTJSUndefined}
		}
		if resolved {
			return moduleVal
		}
		vm.jsThrowReferenceError("Cannot find module '" + moduleName + "'")
		return Value{Type: VTJSUndefined}
	}

	resolved := strings.ToLower(moduleName)
	if strings.HasPrefix(resolved, "node:") {
		resolved = strings.TrimPrefix(resolved, "node:")
	}

	switch resolved {
	case "process":
		return vm.jsNodeGetRootBinding("process")
	case "buffer":
		objID := vm.allocJSID()
		obj := make(map[string]Value, 3)
		obj["__js_type"] = NewString("buffer")
		obj["Buffer"] = vm.jsNodeGetRootBinding("Buffer")
		vm.jsObjectItems[objID] = obj
		vm.jsPropertyItems[objID] = make(map[string]jsPropertyDescriptor, 3)
		return Value{Type: VTJSObject, Num: objID}
	case "path":
		return vm.jsNodeGetRootBinding("path")
	case "os":
		return vm.jsNodeGetRootBinding("os")
	case "fs":
		return vm.jsNodeGetRootBinding("fs")
	case "crypto":
		return vm.jsNodeGetRootBinding("crypto")
	case "http":
		return vm.jsNodeGetRootBinding("http")
	case "https":
		return vm.jsNodeGetRootBinding("https")
	case "querystring":
		return vm.jsNodeGetRootBinding("querystring")
	case "url":
		return vm.jsNodeGetRootBinding("url")
	case "events":
		return vm.jsGetOrCreateEventsModule()
	case "stream":
		return vm.jsGetOrCreateStreamModule()
	default:
		vm.jsThrowReferenceError("Cannot find module '" + moduleName + "'")
		return Value{Type: VTJSUndefined}
	}
}

// jsCreateFSObject allocates the Node.js-compatible fs module object.
func (vm *VM) jsCreateFSObject() Value {
	objID := vm.allocJSID()
	obj := make(map[string]Value, 8)
	obj["__js_type"] = NewString("fs")
	obj["promises"] = vm.jsCreateFSPromisesObject()
	vm.jsObjectItems[objID] = obj
	vm.jsPropertyItems[objID] = make(map[string]jsPropertyDescriptor, 8)
	return Value{Type: VTJSObject, Num: objID}
}

// jsCreateFSPromisesObject allocates the Node.js-compatible fs.promises object.
func (vm *VM) jsCreateFSPromisesObject() Value {
	objID := vm.allocJSID()
	obj := make(map[string]Value, 6)
	obj["__js_type"] = NewString("fs.promises")
	vm.jsObjectItems[objID] = obj
	vm.jsPropertyItems[objID] = make(map[string]jsPropertyDescriptor, 6)
	return Value{Type: VTJSObject, Num: objID}
}

// jsCreateCryptoObject allocates the Node.js-compatible crypto module object.
func (vm *VM) jsCreateCryptoObject() Value {
	objID := vm.allocJSID()
	obj := make(map[string]Value, 8)
	obj["__js_type"] = NewString("crypto")
	vm.jsObjectItems[objID] = obj
	vm.jsPropertyItems[objID] = make(map[string]jsPropertyDescriptor, 8)
	return Value{Type: VTJSObject, Num: objID}
}

// jsCreateHTTPObject allocates the Node.js-compatible http/https module object.
func (vm *VM) jsCreateHTTPObject(moduleType string) Value {
	objID := vm.allocJSID()
	obj := make(map[string]Value, 8)
	obj["__js_type"] = NewString(moduleType)
	vm.jsObjectItems[objID] = obj
	vm.jsPropertyItems[objID] = make(map[string]jsPropertyDescriptor, 8)
	return Value{Type: VTJSObject, Num: objID}
}

// jsNodeGetObjectString returns one string field from a JS object-like value.
func (vm *VM) jsNodeGetObjectString(obj Value, key string) string {
	if obj.Type != VTJSObject && obj.Type != VTJSFunction {
		return ""
	}
	v, deferred := vm.jsMemberGet(obj, key)
	if deferred || v.Type == VTJSUndefined || v.Type == VTNull {
		return ""
	}
	return vm.valueToString(v)
}

// jsNodeGetObjectValue returns one field from a JS object-like value.
func (vm *VM) jsNodeGetObjectValue(obj Value, key string) (Value, bool) {
	if obj.Type != VTJSObject && obj.Type != VTJSFunction {
		return Value{Type: VTJSUndefined}, false
	}
	v, deferred := vm.jsMemberGet(obj, key)
	if deferred || v.Type == VTJSUndefined {
		return Value{Type: VTJSUndefined}, false
	}
	return v, true
}

// jsNodeObjectBoolProperty reads one bool-like property from a JS object.
func (vm *VM) jsNodeObjectBoolProperty(obj Value, key string) bool {
	if obj.Type != VTJSObject && obj.Type != VTJSFunction {
		return false
	}
	v, deferred := vm.jsMemberGet(obj, key)
	if deferred {
		return false
	}
	return vm.jsTruthy(v)
}

// jsNodeValueBytes converts JS data to bytes for fs/crypto/http operations.
func (vm *VM) jsNodeValueBytes(v Value, encoding string) ([]byte, bool) {
	switch v.Type {
	case VTString:
		enc := strings.ToLower(strings.TrimSpace(encoding))
		if enc == "" {
			enc = "utf8"
		}
		if enc == "utf8" || enc == "utf-8" {
			return []byte(v.Str), true
		}
		if enc == "hex" {
			decoded, err := hex.DecodeString(v.Str)
			if err != nil {
				return nil, false
			}
			return decoded, true
		}
		if enc == "base64" {
			decoded, err := base64.StdEncoding.DecodeString(v.Str)
			if err != nil {
				return nil, false
			}
			return decoded, true
		}
		return []byte(v.Str), true
	case VTArray:
		if v.Arr == nil {
			return []byte{}, true
		}
		out := make([]byte, len(v.Arr.Values))
		for i := 0; i < len(v.Arr.Values); i++ {
			out[i] = byte(int(vm.jsToNumber(v.Arr.Values[i]).Flt) & 0xFF)
		}
		return out, true
	case VTJSObject:
		isBuffer := vm.jsObjectStringProperty(v, "__js_type") == "Buffer" ||
			vm.jsObjectStringProperty(v, "__js_class") == "Buffer" ||
			vm.jsObjectStringProperty(v, "__js_ctor") == "Buffer"
		if isBuffer {
			if item, ok := vm.jsBufferItems[v.Num]; ok && item != nil {
				buf := make([]byte, len(item.data))
				copy(buf, item.data)
				return buf, true
			}
			if ref := vm.jsObjectStringProperty(v, "__js_buffer_data"); strings.HasPrefix(ref, "__buffer_") {
				if refID, err := strconv.ParseInt(strings.TrimPrefix(ref, "__buffer_"), 10, 64); err == nil {
					if item, ok := vm.jsBufferItems[refID]; ok && item != nil {
						buf := make([]byte, len(item.data))
						copy(buf, item.data)
						return buf, true
					}
				}
			}
			if utf8Val := vm.jsObjectStringProperty(v, "__js_buffer_utf8"); utf8Val != "" {
				return []byte(utf8Val), true
			}
		}
	}
	strVal := vm.valueToString(v)
	if strings.HasPrefix(strVal, "[JSObject:") && strings.HasSuffix(strVal, "]") {
		idText := strings.TrimSuffix(strings.TrimPrefix(strVal, "[JSObject:"), "]")
		if objID, err := strconv.ParseInt(idText, 10, 64); err == nil {
			if item, ok := vm.jsBufferItems[objID]; ok && item != nil {
				buf := make([]byte, len(item.data))
				copy(buf, item.data)
				return buf, true
			}
		}
	}
	return []byte(strVal), true
}

// jsNodeResolveSandboxPath resolves one fs path inside the server sandbox.
func (vm *VM) jsNodeResolveSandboxPath(v Value) (string, bool) {
	return vm.fsoResolvePath(vm.valueToString(v))
}

// jsNodeExtractEncoding resolves optional encoding argument for fs methods.
func (vm *VM) jsNodeExtractEncoding(args []Value, index int) string {
	if len(args) <= index {
		return ""
	}
	arg := args[index]
	if arg.Type == VTString {
		return strings.ToLower(strings.TrimSpace(arg.Str))
	}
	if arg.Type == VTJSObject || arg.Type == VTJSFunction {
		enc := vm.jsNodeGetObjectString(arg, "encoding")
		return strings.ToLower(strings.TrimSpace(enc))
	}
	return ""
}

// jsCreateFSStatsObject builds one fs.statSync result object.
func (vm *VM) jsCreateFSStatsObject(info os.FileInfo) Value {
	objID := vm.allocJSID()
	obj := make(map[string]Value, 12)
	obj["__js_type"] = NewString("fs.Stats")
	obj["size"] = NewInteger(info.Size())
	obj["mtimeMs"] = NewDouble(float64(info.ModTime().UnixNano()) / 1e6)
	obj["ctimeMs"] = NewDouble(float64(info.ModTime().UnixNano()) / 1e6)
	obj["birthtimeMs"] = NewDouble(float64(info.ModTime().UnixNano()) / 1e6)
	obj["mode"] = NewInteger(int64(info.Mode()))
	obj["_isFile"] = NewBool(!info.IsDir())
	obj["_isDirectory"] = NewBool(info.IsDir())
	vm.jsObjectItems[objID] = obj
	vm.jsPropertyItems[objID] = make(map[string]jsPropertyDescriptor, 12)
	return Value{Type: VTJSObject, Num: objID}
}

// jsNodeCreateDeferredPromise allocates one pending Promise for async fs operations.
func (vm *VM) jsNodeCreateDeferredPromise() Value {
	promiseID := vm.allocJSID()
	promise := Value{Type: VTJSPromise, Num: promiseID}
	vm.jsPromiseItems[promiseID] = &jsPromiseObject{state: jsPromisePending}
	return promise
}

// jsNodeFSBytesToValue converts file bytes to readFile result value based on encoding.
func (vm *VM) jsNodeFSBytesToValue(data []byte, encoding string) (Value, bool) {
	enc := strings.ToLower(strings.TrimSpace(encoding))
	if enc == "" {
		return vm.jsCreateBufferInstance(data), true
	}
	switch enc {
	case "utf8", "utf-8":
		return NewString(string(data)), true
	case "hex":
		return NewString(hex.EncodeToString(data)), true
	case "base64":
		return NewString(base64.StdEncoding.EncodeToString(data)), true
	default:
		return Value{Type: VTJSUndefined}, false
	}
}

// jsQueueAsyncFSRead launches one goroutine to read one sandboxed file path.
func (vm *VM) jsQueueAsyncFSRead(resolvedPath string, encoding string, callback Value, callbackOk bool, promise Value, asPromise bool) bool {
	if vm.jsAsyncFSReadResults == nil {
		return false
	}
	if cap(vm.jsAsyncFSReadResults) == 0 {
		return false
	}
	if len(vm.jsAsyncFSReadResults) >= cap(vm.jsAsyncFSReadResults)-1 {
		return false
	}

	go func(path string, enc string, cb Value, cbOK bool, p Value, usePromise bool, out chan<- jsAsyncFSReadResult) {
		data, err := os.ReadFile(path)
		res := jsAsyncFSReadResult{
			promise:    p,
			callback:   cb,
			data:       data,
			encoding:   enc,
			asPromise:  usePromise,
			callbackOk: cbOK,
		}
		if err != nil {
			res.errMsg = err.Error()
		}
		out <- res
	}(resolvedPath, encoding, callback, callbackOk, promise, asPromise, vm.jsAsyncFSReadResults)

	return true
}

// jsHandleAsyncFSReadResult queues JS-visible callback/promise completion on the microtask queue.
func (vm *VM) jsHandleAsyncFSReadResult(result jsAsyncFSReadResult) {
	vm.jsEnqueueMicrotask(func() {
		if result.asPromise {
			if result.errMsg != "" {
				vm.jsRejectPromise(result.promise, vm.jsCreateErrorObject("Error", result.errMsg))
				return
			}
			val, ok := vm.jsNodeFSBytesToValue(result.data, result.encoding)
			if !ok {
				vm.jsRejectPromise(result.promise, vm.jsCreateErrorObject("TypeError", "Unsupported encoding: "+result.encoding))
				return
			}
			vm.jsResolvePromise(result.promise, val)
			return
		}

		if !result.callbackOk || result.callback.Type != VTJSFunction {
			return
		}

		if result.errMsg != "" {
			errVal := vm.jsCreateErrorObject("Error", result.errMsg)
			vm.jsCall(result.callback, Value{Type: VTJSUndefined}, []Value{errVal, Value{Type: VTJSUndefined}})
			return
		}

		val, ok := vm.jsNodeFSBytesToValue(result.data, result.encoding)
		if !ok {
			errVal := vm.jsCreateErrorObject("TypeError", "Unsupported encoding: "+result.encoding)
			vm.jsCall(result.callback, Value{Type: VTJSUndefined}, []Value{errVal, Value{Type: VTJSUndefined}})
			return
		}

		vm.jsCall(result.callback, Value{Type: VTJSUndefined}, []Value{Value{Type: VTNull}, val})
	})
}

// jsPumpAsyncFSReadResults drains pending fs async read completions into microtasks.
func (vm *VM) jsPumpAsyncFSReadResults(limit int) {
	if vm.jsAsyncFSReadResults == nil || limit <= 0 {
		return
	}
	for i := 0; i < limit; i++ {
		select {
		case result := <-vm.jsAsyncFSReadResults:
			vm.jsHandleAsyncFSReadResult(result)
		default:
			return
		}
	}
}

// jsPumpNodeAsyncTasks drives a single event-loop pass:
//  1. Drain async I/O and timer completions into the microtask queue.
//  2. Run the process.nextTick queue (highest priority).
//  3. Drain the Promise / microtask queue.
//  4. Run setImmediate callbacks.
//
// A re-entrancy guard prevents nested calls (e.g. from within a callback's vm.Run cycle)
// from starting a second overlapping pass.
func (vm *VM) jsPumpNodeAsyncTasks(limit int) {
	if vm.jsPumpingNodeTasks {
		return
	}
	vm.jsPumpingNodeTasks = true
	defer func() { vm.jsPumpingNodeTasks = false }()

	vm.jsPumpAsyncFSReadResults(limit)
	vm.jsPumpTimerResults(limit)
	if len(vm.jsNextTickQueue) > 0 {
		vm.jsProcessNextTickQueue()
	}
	if len(vm.jsMicrotaskQueue) > 0 {
		vm.jsProcessMicrotasks()
	}
	if len(vm.jsImmediateQueue) > 0 {
		vm.jsProcessImmediateQueue()
	}
}

// jsCallFSMethod dispatches fs sync methods.
func (vm *VM) jsCallFSMethod(methodName string, args []Value) (Value, bool) {
	switch strings.ToLower(methodName) {
	case "readfile":
		if len(args) < 2 {
			vm.jsThrowTypeError("fs.readFile requires path and callback")
			return Value{Type: VTJSUndefined}, true
		}
		resolved, ok := vm.jsNodeResolveSandboxPath(args[0])
		if !ok {
			vm.jsThrowTypeError("fs.readFile path is outside sandbox")
			return Value{Type: VTJSUndefined}, true
		}

		callbackIdx := len(args) - 1
		callback := args[callbackIdx]
		if callback.Type != VTJSFunction {
			vm.jsThrowTypeError("fs.readFile callback must be a function")
			return Value{Type: VTJSUndefined}, true
		}

		encoding := ""
		if callbackIdx > 1 {
			encoding = vm.jsNodeExtractEncoding(args, 1)
		}

		if !vm.jsQueueAsyncFSRead(resolved, encoding, callback, true, Value{Type: VTJSUndefined}, false) {
			vm.jsEnqueueMicrotask(func() {
				errVal := vm.jsCreateErrorObject("Error", "fs.readFile async queue is full")
				vm.jsCall(callback, Value{Type: VTJSUndefined}, []Value{errVal, Value{Type: VTJSUndefined}})
			})
		}

		return Value{Type: VTJSUndefined}, true
	case "existssync":
		if len(args) < 1 {
			return NewBool(false), true
		}
		resolved, ok := vm.jsNodeResolveSandboxPath(args[0])
		if !ok {
			return NewBool(false), true
		}
		_, err := os.Stat(resolved)
		return NewBool(err == nil), true
	case "readfilesync":
		if len(args) < 1 {
			vm.jsThrowTypeError("fs.readFileSync requires a path")
			return Value{Type: VTJSUndefined}, true
		}
		resolved, ok := vm.jsNodeResolveSandboxPath(args[0])
		if !ok {
			vm.jsThrowTypeError("fs.readFileSync path is outside sandbox")
			return Value{Type: VTJSUndefined}, true
		}
		data, err := os.ReadFile(resolved)
		if err != nil {
			vm.jsThrowTypeError("fs.readFileSync failed: " + err.Error())
			return Value{Type: VTJSUndefined}, true
		}
		encoding := vm.jsNodeExtractEncoding(args, 1)
		if encoding == "" {
			return vm.jsCreateBufferInstance(data), true
		}
		if encoding == "utf8" || encoding == "utf-8" {
			return NewString(string(data)), true
		}
		if encoding == "hex" {
			return NewString(hex.EncodeToString(data)), true
		}
		if encoding == "base64" {
			return NewString(base64.StdEncoding.EncodeToString(data)), true
		}
		vm.jsThrowTypeError("Unsupported encoding: " + encoding)
		return Value{Type: VTJSUndefined}, true
	case "writefilesync":
		if len(args) < 2 {
			vm.jsThrowTypeError("fs.writeFileSync requires path and data")
			return Value{Type: VTJSUndefined}, true
		}
		resolved, ok := vm.jsNodeResolveSandboxPath(args[0])
		if !ok {
			vm.jsThrowTypeError("fs.writeFileSync path is outside sandbox")
			return Value{Type: VTJSUndefined}, true
		}
		encoding := vm.jsNodeExtractEncoding(args, 2)
		data, bytesOK := vm.jsNodeValueBytes(args[1], encoding)
		if !bytesOK {
			vm.jsThrowTypeError("fs.writeFileSync failed to convert data")
			return Value{Type: VTJSUndefined}, true
		}
		if err := os.MkdirAll(filepath.Dir(resolved), 0o755); err != nil {
			vm.jsThrowTypeError("fs.writeFileSync mkdir failed: " + err.Error())
			return Value{Type: VTJSUndefined}, true
		}
		if err := os.WriteFile(resolved, data, 0o644); err != nil {
			vm.jsThrowTypeError("fs.writeFileSync failed: " + err.Error())
			return Value{Type: VTJSUndefined}, true
		}
		return Value{Type: VTJSUndefined}, true
	case "statsync":
		if len(args) < 1 {
			vm.jsThrowTypeError("fs.statSync requires a path")
			return Value{Type: VTJSUndefined}, true
		}
		resolved, ok := vm.jsNodeResolveSandboxPath(args[0])
		if !ok {
			vm.jsThrowTypeError("fs.statSync path is outside sandbox")
			return Value{Type: VTJSUndefined}, true
		}
		info, err := os.Stat(resolved)
		if err != nil {
			vm.jsThrowTypeError("fs.statSync failed: " + err.Error())
			return Value{Type: VTJSUndefined}, true
		}
		return vm.jsCreateFSStatsObject(info), true
	}
	return Value{Type: VTJSUndefined}, false
}

// jsCallFSPromisesMethod dispatches fs.promises API methods.
func (vm *VM) jsCallFSPromisesMethod(methodName string, args []Value) (Value, bool) {
	switch strings.ToLower(methodName) {
	case "readfile":
		if len(args) < 1 {
			vm.jsThrowTypeError("fs.promises.readFile requires a path")
			return Value{Type: VTJSUndefined}, true
		}
		resolved, ok := vm.jsNodeResolveSandboxPath(args[0])
		if !ok {
			vm.jsThrowTypeError("fs.promises.readFile path is outside sandbox")
			return Value{Type: VTJSUndefined}, true
		}
		encoding := vm.jsNodeExtractEncoding(args, 1)
		promise := vm.jsNodeCreateDeferredPromise()
		if !vm.jsQueueAsyncFSRead(resolved, encoding, Value{Type: VTJSUndefined}, false, promise, true) {
			vm.jsRejectPromise(promise, vm.jsCreateErrorObject("Error", "fs.promises.readFile async queue is full"))
		}
		return promise, true
	}
	return Value{Type: VTJSUndefined}, false
}

// jsCallFSStatsMethod dispatches fs.Stats instance methods.
func (vm *VM) jsCallFSStatsMethod(target Value, methodName string) (Value, bool) {
	if target.Type != VTJSObject {
		return Value{Type: VTJSUndefined}, false
	}
	switch strings.ToLower(methodName) {
	case "isfile":
		return NewBool(vm.jsNodeObjectBoolProperty(target, "_isFile")), true
	case "isdirectory":
		return NewBool(vm.jsNodeObjectBoolProperty(target, "_isDirectory")), true
	}
	return Value{Type: VTJSUndefined}, false
}

// jsCreateNodeHashStateObject allocates one crypto hash/hmac state object.
func (vm *VM) jsCreateNodeHashStateObject(typeName string, algorithm string, key []byte) Value {
	objID := vm.allocJSID()
	obj := make(map[string]Value, 10)
	obj["__js_type"] = NewString(typeName)
	obj["__js_crypto_algorithm"] = NewString(strings.ToLower(strings.TrimSpace(algorithm)))
	obj["__js_crypto_data"] = NewString("")
	obj["__js_crypto_finalized"] = NewBool(false)
	obj["__js_crypto_key"] = vm.jsCreateBufferInstance(key)
	vm.jsObjectItems[objID] = obj
	vm.jsPropertyItems[objID] = make(map[string]jsPropertyDescriptor, 10)
	return Value{Type: VTJSObject, Num: objID}
}

// jsNodeComputeDigest computes one digest for hash or hmac objects.
func (vm *VM) jsNodeComputeDigest(target Value) ([]byte, bool) {
	obj := vm.jsObjectItems[target.Num]
	if obj == nil {
		return nil, false
	}
	alg := strings.ToLower(strings.TrimSpace(vm.jsObjectStringProperty(target, "__js_crypto_algorithm")))
	data := []byte(vm.jsObjectStringProperty(target, "__js_crypto_data"))
	typeName := vm.jsObjectStringProperty(target, "__js_type")
	if typeName == "crypto.Hmac" {
		keyVal := obj["__js_crypto_key"]
		key, _ := vm.jsNodeValueBytes(keyVal, "")
		switch alg {
		case "sha1":
			h := hmac.New(sha1.New, key)
			_, _ = h.Write(data)
			return h.Sum(nil), true
		case "sha256", "":
			h := hmac.New(sha256.New, key)
			_, _ = h.Write(data)
			return h.Sum(nil), true
		case "md5":
			h := hmac.New(md5.New, key)
			_, _ = h.Write(data)
			return h.Sum(nil), true
		default:
			return nil, false
		}
	}
	switch alg {
	case "md5":
		sum := md5.Sum(data)
		return sum[:], true
	case "sha1":
		sum := sha1.Sum(data)
		return sum[:], true
	case "sha256", "":
		sum := sha256.Sum256(data)
		return sum[:], true
	default:
		return nil, false
	}
}

// jsCallCryptoMethod dispatches top-level crypto module methods.
func (vm *VM) jsCallCryptoMethod(methodName string, args []Value) (Value, bool) {
	switch strings.ToLower(methodName) {
	case "createhash":
		if len(args) < 1 {
			vm.jsThrowTypeError("crypto.createHash requires an algorithm")
			return Value{Type: VTJSUndefined}, true
		}
		alg := strings.ToLower(strings.TrimSpace(vm.valueToString(args[0])))
		if alg != "md5" && alg != "sha1" && alg != "sha256" {
			vm.jsThrowTypeError("Unsupported hash algorithm: " + alg)
			return Value{Type: VTJSUndefined}, true
		}
		return vm.jsCreateNodeHashStateObject("crypto.Hash", alg, nil), true
	case "createhmac":
		if len(args) < 2 {
			vm.jsThrowTypeError("crypto.createHmac requires algorithm and key")
			return Value{Type: VTJSUndefined}, true
		}
		alg := strings.ToLower(strings.TrimSpace(vm.valueToString(args[0])))
		if alg != "md5" && alg != "sha1" && alg != "sha256" {
			vm.jsThrowTypeError("Unsupported hmac algorithm: " + alg)
			return Value{Type: VTJSUndefined}, true
		}
		key, ok := vm.jsNodeValueBytes(args[1], "")
		if !ok {
			vm.jsThrowTypeError("Invalid HMAC key")
			return Value{Type: VTJSUndefined}, true
		}
		return vm.jsCreateNodeHashStateObject("crypto.Hmac", alg, key), true
	case "randombytes":
		if len(args) < 1 {
			vm.jsThrowTypeError("crypto.randomBytes requires size")
			return Value{Type: VTJSUndefined}, true
		}
		size := int(vm.jsToNumber(args[0]).Flt)
		if size < 0 {
			vm.jsThrowRangeError("randomBytes size must be non-negative")
			return Value{Type: VTJSUndefined}, true
		}
		buf := make([]byte, size)
		if _, err := rand.Read(buf); err != nil {
			vm.jsThrowTypeError("crypto.randomBytes failed: " + err.Error())
			return Value{Type: VTJSUndefined}, true
		}
		return vm.jsCreateBufferInstance(buf), true
	}
	return Value{Type: VTJSUndefined}, false
}

// jsCallCryptoHashMethod dispatches methods shared by hash and hmac state objects.
func (vm *VM) jsCallCryptoHashMethod(target Value, methodName string, args []Value) (Value, bool) {
	if target.Type != VTJSObject {
		return Value{Type: VTJSUndefined}, false
	}
	obj := vm.jsObjectItems[target.Num]
	if obj == nil {
		return Value{Type: VTJSUndefined}, false
	}
	switch strings.ToLower(methodName) {
	case "update":
		if vm.jsNodeObjectBoolProperty(target, "__js_crypto_finalized") {
			vm.jsThrowTypeError("Digest already called")
			return Value{Type: VTJSUndefined}, true
		}
		if len(args) < 1 {
			vm.jsThrowTypeError("update requires data")
			return Value{Type: VTJSUndefined}, true
		}
		encoding := ""
		if len(args) > 1 {
			encoding = vm.valueToString(args[1])
		}
		data, ok := vm.jsNodeValueBytes(args[0], encoding)
		if !ok {
			vm.jsThrowTypeError("Invalid update payload")
			return Value{Type: VTJSUndefined}, true
		}
		obj["__js_crypto_data"] = NewString(vm.jsObjectStringProperty(target, "__js_crypto_data") + string(data))
		return target, true
	case "digest":
		if vm.jsNodeObjectBoolProperty(target, "__js_crypto_finalized") {
			vm.jsThrowTypeError("Digest already called")
			return Value{Type: VTJSUndefined}, true
		}
		digest, ok := vm.jsNodeComputeDigest(target)
		if !ok {
			vm.jsThrowTypeError("Unsupported digest algorithm")
			return Value{Type: VTJSUndefined}, true
		}
		obj["__js_crypto_finalized"] = NewBool(true)
		enc := ""
		if len(args) > 0 {
			enc = strings.ToLower(strings.TrimSpace(vm.valueToString(args[0])))
		}
		if enc == "" {
			return vm.jsCreateBufferInstance(digest), true
		}
		if enc == "hex" {
			return NewString(hex.EncodeToString(digest)), true
		}
		if enc == "base64" {
			return NewString(base64.StdEncoding.EncodeToString(digest)), true
		}
		if enc == "latin1" {
			return NewString(string(digest)), true
		}
		vm.jsThrowTypeError("Unsupported digest encoding: " + enc)
		return Value{Type: VTJSUndefined}, true
	}
	return Value{Type: VTJSUndefined}, false
}

// jsNodeReadHeadersObject converts a JS headers object to an HTTP header map.
func (vm *VM) jsNodeReadHeadersObject(v Value) http.Header {
	headers := make(http.Header)
	if v.Type != VTJSObject && v.Type != VTJSFunction {
		return headers
	}
	keys := vm.jsObjectOwnPropertyNames(v)
	for i := 0; i < len(keys); i++ {
		k := keys[i]
		if strings.HasPrefix(k, "__js_") {
			continue
		}
		val, deferred := vm.jsMemberGet(v, k)
		if deferred || val.Type == VTJSUndefined || val.Type == VTNull {
			continue
		}
		headers.Set(k, vm.valueToString(val))
	}
	return headers
}

// jsNodeBuildRequestURL builds a request URL string from Node-like options.
func (vm *VM) jsNodeBuildRequestURL(moduleType string, options Value) string {
	if options.Type == VTString {
		return options.Str
	}
	urlStr := vm.jsNodeGetObjectString(options, "url")
	if urlStr == "" {
		urlStr = vm.jsNodeGetObjectString(options, "href")
	}
	if urlStr != "" {
		if strings.HasPrefix(urlStr, "//") {
			if moduleType == "https" {
				return "https:" + urlStr
			}
			return "http:" + urlStr
		}
		return urlStr
	}

	host := vm.jsNodeGetObjectString(options, "hostname")
	if host == "" {
		host = vm.jsNodeGetObjectString(options, "host")
	}
	if host == "" {
		return ""
	}
	port := vm.jsNodeGetObjectString(options, "port")
	path := vm.jsNodeGetObjectString(options, "path")
	if path == "" {
		path = "/"
	}
	protocol := vm.jsNodeGetObjectString(options, "protocol")
	if protocol == "" {
		if moduleType == "https" {
			protocol = "https:"
		} else {
			protocol = "http:"
		}
	}
	protocol = strings.TrimSuffix(protocol, ":")
	if port != "" {
		host = host + ":" + port
	}
	return protocol + "://" + host + path
}

// jsCreateHTTPResponseObject creates a Node-like incoming message object.
func (vm *VM) jsCreateHTTPResponseObject(resp *http.Response, body []byte) Value {
	headersID := vm.allocJSID()
	headersObj := make(map[string]Value, len(resp.Header)+2)
	headersObj["__js_type"] = NewString("Object")
	for key, values := range resp.Header {
		if len(values) == 0 {
			headersObj[key] = NewString("")
			continue
		}
		headersObj[key] = NewString(values[0])
	}
	vm.jsObjectItems[headersID] = headersObj
	vm.jsPropertyItems[headersID] = make(map[string]jsPropertyDescriptor, len(headersObj)+2)

	objID := vm.allocJSID()
	obj := make(map[string]Value, 12)
	obj["__js_type"] = NewString("http.IncomingMessage")
	obj["statusCode"] = NewInteger(int64(resp.StatusCode))
	obj["statusMessage"] = NewString(resp.Status)
	obj["headers"] = Value{Type: VTJSObject, Num: headersID}
	obj["ok"] = NewBool(resp.StatusCode >= 200 && resp.StatusCode < 300)
	obj["body"] = NewString(string(body))
	vm.jsObjectItems[objID] = obj
	vm.jsPropertyItems[objID] = make(map[string]jsPropertyDescriptor, 12)
	return Value{Type: VTJSObject, Num: objID}
}

// jsNodeDoHTTPRequest executes one HTTP(S) request and returns a response object.
func (vm *VM) jsNodeDoHTTPRequest(moduleType string, method string, args []Value) (Value, bool) {
	if len(args) < 1 {
		vm.jsThrowTypeError(moduleType + ".request requires URL or options")
		return Value{Type: VTJSUndefined}, true
	}

	target := args[0]
	urlStr := vm.jsNodeBuildRequestURL(moduleType, target)
	if urlStr == "" {
		vm.jsThrowTypeError(moduleType + ".request missing URL")
		return Value{Type: VTJSUndefined}, true
	}

	options := Value{Type: VTJSUndefined}
	if target.Type == VTJSObject || target.Type == VTJSFunction {
		options = target
	} else if len(args) > 1 && (args[1].Type == VTJSObject || args[1].Type == VTJSFunction) {
		options = args[1]
	}

	reqMethod := strings.ToUpper(strings.TrimSpace(method))
	if reqMethod == "" {
		reqMethod = strings.ToUpper(strings.TrimSpace(vm.jsNodeGetObjectString(options, "method")))
	}
	if reqMethod == "" {
		reqMethod = "GET"
	}

	bodyBytes := []byte(nil)
	if payload, ok := vm.jsNodeGetObjectValue(options, "body"); ok {
		data, bytesOK := vm.jsNodeValueBytes(payload, "")
		if bytesOK {
			bodyBytes = data
		}
	} else if len(args) > 1 && args[1].Type != VTJSObject && args[1].Type != VTJSFunction {
		data, bytesOK := vm.jsNodeValueBytes(args[1], "")
		if bytesOK {
			bodyBytes = data
		}
	}

	var bodyReader io.Reader
	if len(bodyBytes) > 0 {
		bodyReader = bytes.NewReader(bodyBytes)
	}

	req, err := http.NewRequest(reqMethod, urlStr, bodyReader)
	if err != nil {
		vm.jsThrowTypeError(moduleType + ".request failed: " + err.Error())
		return Value{Type: VTJSUndefined}, true
	}

	if options.Type == VTJSObject || options.Type == VTJSFunction {
		headersVal, ok := vm.jsNodeGetObjectValue(options, "headers")
		if ok {
			headers := vm.jsNodeReadHeadersObject(headersVal)
			for key, vals := range headers {
				for i := 0; i < len(vals); i++ {
					req.Header.Add(key, vals[i])
				}
			}
		}
	}

	if len(bodyBytes) > 0 && req.Header.Get("Content-Type") == "" {
		req.Header.Set("Content-Type", "application/octet-stream")
	}

	timeout := 10 * time.Second
	if options.Type == VTJSObject || options.Type == VTJSFunction {
		timeoutRaw := vm.jsNodeGetObjectString(options, "timeout")
		if timeoutRaw != "" {
			ms, parseErr := strconv.Atoi(timeoutRaw)
			if parseErr == nil && ms > 0 {
				timeout = time.Duration(ms) * time.Millisecond
			}
		}
	}

	client := &http.Client{Timeout: timeout}
	resp, err := client.Do(req)
	if err != nil {
		vm.jsThrowTypeError(moduleType + ".request failed: " + err.Error())
		return Value{Type: VTJSUndefined}, true
	}
	defer resp.Body.Close()

	respBody, err := io.ReadAll(resp.Body)
	if err != nil {
		vm.jsThrowTypeError(moduleType + ".request read failed: " + err.Error())
		return Value{Type: VTJSUndefined}, true
	}

	return vm.jsCreateHTTPResponseObject(resp, respBody), true
}

// jsCallHTTPMethod dispatches http/https module methods.
func (vm *VM) jsCallHTTPMethod(moduleType string, methodName string, args []Value) (Value, bool) {
	switch strings.ToLower(methodName) {
	case "get":
		return vm.jsNodeDoHTTPRequest(moduleType, "GET", args)
	case "request":
		return vm.jsNodeDoHTTPRequest(moduleType, "", args)
	}
	return Value{Type: VTJSUndefined}, false
}

// jsCallHTTPResponseMethod dispatches incoming-message helper methods.
func (vm *VM) jsCallHTTPResponseMethod(target Value, methodName string, _ []Value) (Value, bool) {
	if target.Type != VTJSObject {
		return Value{Type: VTJSUndefined}, false
	}
	switch strings.ToLower(methodName) {
	case "text":
		return NewString(vm.jsObjectStringProperty(target, "body")), true
	case "json":
		jsonLib := &G3JSON{vm: vm}
		parsed := jsonLib.DispatchMethod("Parse", []Value{NewString(vm.jsObjectStringProperty(target, "body"))})
		if parsed.Type == VTEmpty {
			return Value{Type: VTNull}, true
		}
		return parsed, true
	}
	return Value{Type: VTJSUndefined}, false
}
