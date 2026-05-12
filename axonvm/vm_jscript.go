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
	"encoding/json"
	"fmt"
	"math"
	"math/big"
	"math/rand"
	"path/filepath"
	"regexp"
	"sort"
	"strconv"
	"strings"
	"time"

	"g3pix.com.br/axonasp/jscript/ftoa"
	jsparser "g3pix.com.br/axonasp/jscript/parser"
	"g3pix.com.br/axonasp/vbscript"
)

const jsMaxStringBytes = 8 * 1024 * 1024
const jsMaxStringWorkBytes = 2 * 1024 * 1024
const jsInternalPropPrefix = "__js_"
const jsAccessorGetterPrefix = "__js_getter__"
const jsAccessorSetterPrefix = "__js_setter__"
const jsSymbolPropertyPrefix = "__js_sym__"
const jsRestParamPrefix = "__js_rest__:"
const jsClassConstructorFlag = "__axon_internal__:class_constructor"
const jsStrictModeFlag = "__axon_internal__:strict_mode"
const jsGeneratorFlag = "__axon_internal__:generator"
const jsAsyncFlag = "__axon_internal__:async"
const jsDerivedConstructorFlag = "__axon_internal__:derived_constructor"
const jsModuleExportPrefix = "__js_export__:"

const (
	jsPropertyKindMethod = 0
	jsPropertyKindGet    = 1
	jsPropertyKindSet    = 2
)

type jsObjectState struct {
	Extensible bool
}

type jsPropertyDescriptor struct {
	Value        Value
	Getter       Value
	Setter       Value
	HasValue     bool
	HasGetter    bool
	HasSetter    bool
	Enumerable   bool
	Configurable bool
	Writable     bool
}

type jsEnvFrame struct {
	parentID int64
	bindings map[string]Value
}

type jsArgumentsBinding struct {
	envID        int64
	indexToParam map[string]string
	paramToIndex map[string]string
}

type jsDefinePropertySpec struct {
	desc            jsPropertyDescriptor
	hasEnumerable   bool
	hasConfigurable bool
	hasWritable     bool
}

type jsFunctionObject struct {
	name       string
	params     []string
	restParam  string
	localCount int
	startIP    int
	endIP      int
	envID      int64
	protoID    int64
	isBound    bool
	boundFn    Value
	boundThis  Value
	boundArgs  []Value
	// isArrow marks this as an ES6 arrow function that captures 'this' lexically.
	// When true, capturedThis is used as the receiver regardless of call site.
	isArrow            bool
	capturedThis       Value
	isClassConstructor bool
	isStrict           bool
	isDerived          bool
	homeObjID          int64
	isAsync            bool
	isGenerator        bool
}

type jsCallFrame struct {
	returnIP     int
	envID        int64
	savedFP      int
	fn           Value
	thisVal      Value
	newTarget    Value
	tryDepth     int
	savedSP      int
	isCtor       bool
	ctorObj      Value
	jsStrictMode bool
	isSuperCall  bool
}

type jsForInEnumerator struct {
	keys  []string
	index int
}

// jsForOfEnumerator holds the collected values for a for...of loop.
// Values are collected once when the loop is first entered and then
// consumed one per iteration to avoid mutating the source during traversal.
type jsForOfEnumerator struct {
	values []Value
	index  int
}

type jsPromiseState uint8

const (
	jsPromisePending jsPromiseState = iota
	jsPromiseFulfilled
	jsPromiseRejected
)

type jsPromiseReaction struct {
	onFulfilled Value
	onRejected  Value
	capability  *jsPromiseCapability
}

type jsPromiseCapability struct {
	promise Value
	resolve Value
	reject  Value
}

type jsPromiseObject struct {
	state     jsPromiseState
	result    Value
	reactions []jsPromiseReaction
	handled   bool // true if a rejection handler was attached
}

func (vm *VM) jsModuleExportKey(name string) string {
	return jsModuleExportPrefix + name
}

func (vm *VM) jsSetModuleExport(name string, value Value) {
	vm.ensureJSRootEnv()
	env := vm.jsEnvItems[vm.jsActiveEnvID]
	if env == nil {
		env = &jsEnvFrame{bindings: make(map[string]Value, 8)}
		vm.jsEnvItems[vm.jsActiveEnvID] = env
	}
	env.bindings[vm.jsModuleExportKey(name)] = value
}

func (vm *VM) jsResolveModulePath(specifier string) (string, error) {
	trimmed := strings.TrimSpace(specifier)
	if trimmed == "" {
		return "", fmt.Errorf("empty module specifier")
	}

	resolved := trimmed
	if !filepath.IsAbs(resolved) {
		baseFile := strings.TrimSpace(vm.sourceName)
		if baseFile == "" {
			baseFile = strings.TrimSpace(vm.baseSourceName)
		}
		if baseFile == "" {
			baseFile = "."
		}
		resolved = filepath.Join(filepath.Dir(baseFile), resolved)
	}
	if filepath.Ext(resolved) == "" {
		resolved += ".js"
	}
	absPath, err := filepath.Abs(resolved)
	if err != nil {
		return "", err
	}
	return absPath, nil
}

func (vm *VM) jsImportModule(specifier string) (*jsEnvFrame, bool) {
	modulePath, err := vm.jsResolveModulePath(specifier)
	if err != nil {
		vm.jsThrowReferenceError("Cannot resolve module '" + specifier + "': " + err.Error())
		return nil, false
	}

	if env, ok := vm.jsModuleInstances[modulePath]; ok && env != nil {
		return env, true
	}

	moduleEnvID := vm.allocJSID()
	moduleEnv := &jsEnvFrame{parentID: 0, bindings: make(map[string]Value, 16)}
	vm.jsEnvItems[moduleEnvID] = moduleEnv
	vm.jsModuleInstances[modulePath] = moduleEnv

	cache := getExecuteScriptCache()
	program, loadErr := cache.LoadOrCompileWithMode(modulePath, vm.executionMode)
	if loadErr != nil {
		delete(vm.jsModuleInstances, modulePath)
		delete(vm.jsEnvItems, moduleEnvID)
		vm.jsThrowReferenceError("Cannot load module '" + modulePath + "': " + loadErr.Error())
		return nil, false
	}

	startIP := vm.appendExecuteProgram(program.GlobalCount, program.Constants, program.Bytecode)
	child := vm.cloneForExecuteLocal(startIP)
	child.sourceName = modulePath
	child.baseSourceName = modulePath
	child.jsActiveEnvID = moduleEnvID
	if runErr := child.Run(); runErr != nil {
		delete(vm.jsModuleInstances, modulePath)
		delete(vm.jsEnvItems, moduleEnvID)
		vm.jsThrowReferenceError("Error executing module '" + modulePath + "': " + runErr.Error())
		return nil, false
	}

	vm.syncExecuteGlobalState(child)
	if finalEnv, ok := vm.jsModuleInstances[modulePath]; ok && finalEnv != nil {
		return finalEnv, true
	}
	return moduleEnv, true
}

func (vm *VM) jsEval(args []Value) Value {
	if len(args) == 0 {
		return Value{Type: VTJSUndefined}
	}

	if args[0].Type != VTString {
		return args[0]
	}

	expr := strings.TrimSpace(args[0].String())
	expr = strings.TrimLeft(expr, "\uFEFF")
	if expr == "" {
		return Value{Type: VTJSUndefined}
	}

	compiler := NewASPCompiler("")
	compiler.sourceName = vm.sourceName
	compiler.compileJScriptEvalSnippet(expr)

	if len(compiler.bytecode) == 0 {
		return Value{Type: VTJSUndefined}
	}

	startIP := vm.appendExecuteProgram(compiler.GlobalsCount(), compiler.constants, compiler.bytecode)
	if startIP < 0 || startIP >= len(vm.bytecode) {
		return Value{Type: VTJSUndefined}
	}

	child := vm.cloneForExecuteLocal(startIP)
	if err := child.Run(); err != nil {
		vm.syncExecuteGlobalState(child)
		return Value{Type: VTJSUndefined}
	}

	resultValue := Value{Type: VTJSUndefined}
	if child.sp >= 0 {
		resultValue = child.stack[child.sp]
	}

	vm.syncExecuteGlobalState(child)
	return resultValue
}

func defaultJSObjectState() jsObjectState {
	return jsObjectState{Extensible: true}
}

func (vm *VM) jsGetObjectState(objID int64) jsObjectState {
	state, ok := vm.jsObjectStateItems[objID]
	if ok {
		return state
	}
	return defaultJSObjectState()
}

func (vm *VM) jsSetObjectExtensible(objID int64, extensible bool) {
	state := vm.jsGetObjectState(objID)
	state.Extensible = extensible
	vm.jsObjectStateItems[objID] = state
}

func (vm *VM) jsObjectOwnPropertyNames(target Value) []string {
	if target.Type != VTJSObject && target.Type != VTJSFunction {
		return nil
	}
	keys := make(map[string]struct{}, 8)
	if obj, ok := vm.jsObjectItems[target.Num]; ok {
		for key := range obj {
			if strings.HasPrefix(key, jsInternalPropPrefix) {
				continue
			}
			if strings.HasPrefix(key, jsSymbolPropertyPrefix) {
				continue
			}
			keys[key] = struct{}{}
		}
	}
	if props, ok := vm.jsPropertyItems[target.Num]; ok {
		for key := range props {
			if strings.HasPrefix(key, jsInternalPropPrefix) {
				continue
			}
			if strings.HasPrefix(key, jsSymbolPropertyPrefix) {
				continue
			}
			keys[key] = struct{}{}
		}
	}
	if len(keys) == 0 {
		return nil
	}
	out := make([]string, 0, len(keys))
	for key := range keys {
		out = append(out, key)
	}
	sort.Strings(out)
	return out
}

func (vm *VM) jsObjectIsExtensible(target Value) bool {
	if target.Type != VTJSObject && target.Type != VTJSFunction {
		return false
	}
	return vm.jsGetObjectState(target.Num).Extensible
}

func (vm *VM) jsObjectSeal(target Value) Value {
	if target.Type != VTJSObject && target.Type != VTJSFunction {
		return target
	}
	names := vm.jsObjectOwnPropertyNames(target)
	for i := 0; i < len(names); i++ {
		desc, ok := vm.jsGetDescriptor(target.Num, names[i])
		if !ok {
			continue
		}
		desc.Configurable = false
		vm.jsSetDescriptor(target.Num, names[i], desc)
	}
	vm.jsSetObjectExtensible(target.Num, false)
	return target
}

func (vm *VM) jsObjectFreeze(target Value) Value {
	if target.Type != VTJSObject && target.Type != VTJSFunction {
		return target
	}
	names := vm.jsObjectOwnPropertyNames(target)
	for i := 0; i < len(names); i++ {
		desc, ok := vm.jsGetDescriptor(target.Num, names[i])
		if !ok {
			continue
		}
		desc.Configurable = false
		if desc.HasValue {
			desc.Writable = false
		}
		vm.jsSetDescriptor(target.Num, names[i], desc)
	}
	vm.jsSetObjectExtensible(target.Num, false)
	return target
}

func (vm *VM) jsObjectIsSealed(target Value) bool {
	if target.Type != VTJSObject && target.Type != VTJSFunction {
		return false
	}
	if vm.jsObjectIsExtensible(target) {
		return false
	}
	names := vm.jsObjectOwnPropertyNames(target)
	for i := 0; i < len(names); i++ {
		desc, ok := vm.jsGetDescriptor(target.Num, names[i])
		if ok && desc.Configurable {
			return false
		}
	}
	return true
}

func (vm *VM) jsObjectIsFrozen(target Value) bool {
	if !vm.jsObjectIsSealed(target) {
		return false
	}
	names := vm.jsObjectOwnPropertyNames(target)
	for i := 0; i < len(names); i++ {
		desc, ok := vm.jsGetDescriptor(target.Num, names[i])
		if ok && desc.HasValue && desc.Writable {
			return false
		}
	}
	return true
}

func (vm *VM) jsEnsurePropertyMap(objID int64) map[string]jsPropertyDescriptor {
	props, ok := vm.jsPropertyItems[objID]
	if ok {
		return props
	}
	props = make(map[string]jsPropertyDescriptor, 8)
	vm.jsPropertyItems[objID] = props
	return props
}

func jsDefaultPropertyDescriptor(v Value) jsPropertyDescriptor {
	return jsPropertyDescriptor{Value: v, HasValue: true, Enumerable: true, Configurable: true, Writable: true}
}

func (vm *VM) jsSetProto(target Value, proto Value) {
	if target.Type != VTJSObject && target.Type != VTJSFunction {
		return
	}
	id := target.Num
	obj := vm.jsObjectItems[id]
	if obj == nil {
		obj = make(map[string]Value, 2)
		vm.jsObjectItems[id] = obj
	}
	obj["__js_proto"] = proto
	vm.jsInvalidateObjectIC(id)
}

// jsInvalidateObjectIC discards one object's slot/layout metadata.
func (vm *VM) jsInvalidateObjectIC(objID int64) {
	delete(vm.jsObjectShape, objID)
	delete(vm.jsObjectSlots, objID)
	delete(vm.jsObjectSlotIndex, objID)
}

func (vm *VM) jsEnsureObjectICLayout(objID int64) bool {
	obj, ok := vm.jsObjectItems[objID]
	if !ok {
		return false
	}
	keys := make([]string, 0, len(obj))
	for key := range obj {
		if strings.HasPrefix(key, jsInternalPropPrefix) {
			continue
		}
		keys = append(keys, key)
	}
	sort.Strings(keys)

	var shapeSig strings.Builder
	for i := 0; i < len(keys); i++ {
		if i > 0 {
			shapeSig.WriteByte('\x1f')
		}
		shapeSig.WriteString(keys[i])
	}

	shapeID := vm.jsShapeBySignature[shapeSig.String()]
	if shapeID == 0 {
		shapeID = vm.jsNextShapeID
		if shapeID == 0 {
			shapeID = 1
		}
		vm.jsNextShapeID = shapeID + 1
		vm.jsShapeBySignature[shapeSig.String()] = shapeID
		if len(keys) > 0 {
			layout := make([]string, len(keys))
			copy(layout, keys)
			vm.jsShapeSlots[shapeID] = layout
		}
	}

	slots := make([]Value, len(keys))
	indexByName := make(map[string]uint16, len(keys))
	for i := 0; i < len(keys); i++ {
		k := keys[i]
		slots[i] = obj[k]
		indexByName[k] = uint16(i)
	}

	vm.jsObjectShape[objID] = shapeID
	vm.jsObjectSlots[objID] = slots
	vm.jsObjectSlotIndex[objID] = indexByName
	return true
}

func (vm *VM) jsResolveICSlot(target Value, member string) (uint32, uint16, bool) {
	if target.Type != VTJSObject && target.Type != VTJSFunction {
		return 0, 0, false
	}
	if !vm.jsEnsureObjectICLayout(target.Num) {
		return 0, 0, false
	}
	shapeID := vm.jsObjectShape[target.Num]
	if shapeID == 0 {
		return 0, 0, false
	}
	slot, ok := vm.jsObjectSlotIndex[target.Num][member]
	if !ok {
		return 0, 0, false
	}
	desc, hasDesc := vm.jsGetDescriptor(target.Num, member)
	if hasDesc {
		if desc.HasGetter || desc.HasSetter || !desc.HasValue {
			return 0, 0, false
		}
	}
	return shapeID, slot, true
}

func (vm *VM) jsICMemberGet(target Value, member string, shapeID uint32, slot uint16, flags uint16) (Value, bool) {
	if flags == 0 || shapeID == 0 {
		return Value{Type: VTJSUndefined}, false
	}
	if target.Type != VTJSObject && target.Type != VTJSFunction {
		return Value{Type: VTJSUndefined}, false
	}
	if vm.jsObjectShape[target.Num] != shapeID {
		return Value{Type: VTJSUndefined}, false
	}
	layout := vm.jsShapeSlots[shapeID]
	if int(slot) >= len(layout) || layout[slot] != member {
		return Value{Type: VTJSUndefined}, false
	}
	slots := vm.jsObjectSlots[target.Num]
	if int(slot) >= len(slots) {
		return Value{Type: VTJSUndefined}, false
	}
	return slots[slot], true
}

func (vm *VM) jsICMemberSet(target Value, member string, val Value, shapeID uint32, slot uint16, flags uint16) bool {
	if flags == 0 || shapeID == 0 {
		return false
	}
	if target.Type != VTJSObject && target.Type != VTJSFunction {
		return false
	}
	if vm.jsObjectShape[target.Num] != shapeID {
		return false
	}
	layout := vm.jsShapeSlots[shapeID]
	if int(slot) >= len(layout) || layout[slot] != member {
		return false
	}
	desc, hasDesc := vm.jsGetDescriptor(target.Num, member)
	if hasDesc {
		if desc.HasSetter || !desc.Writable || !desc.HasValue {
			return false
		}
	}
	obj := vm.jsObjectItems[target.Num]
	if obj == nil {
		return false
	}
	obj[member] = val
	slots := vm.jsObjectSlots[target.Num]
	if int(slot) >= len(slots) {
		return false
	}
	slots[slot] = val
	vm.jsObjectSlots[target.Num] = slots
	if hasDesc {
		desc.Value = val
		desc.HasValue = true
		vm.jsPropertyItems[target.Num][member] = desc
	}
	return true
}

func (vm *VM) jsICPopulate(cachePos int, target Value, member string) {
	shapeID, slot, ok := vm.jsResolveICSlot(target, member)
	if !ok || cachePos+8 > len(vm.bytecode) {
		return
	}
	binary.BigEndian.PutUint32(vm.bytecode[cachePos:], shapeID)
	binary.BigEndian.PutUint16(vm.bytecode[cachePos+4:], slot)
	binary.BigEndian.PutUint16(vm.bytecode[cachePos+6:], 1)
}

// jsClassInherit wires a derived class to its superclass and returns the subclass.
func (vm *VM) jsClassInherit(subclass Value, superclass Value) Value {
	if subclass.Type != VTJSFunction && subclass.Type != VTJSObject {
		vm.jsThrowTypeError("Class constructor cannot be used as a base value")
		return Value{Type: VTJSUndefined}
	}

	if superclass.Type == VTNull {
		proto, deferred := vm.jsMemberGet(subclass, "prototype")
		if !deferred && proto.Type == VTJSObject {
			vm.jsSetProto(proto, NewNull())
		}
		return subclass
	}

	if superclass.Type != VTJSFunction && superclass.Type != VTJSObject {
		vm.jsThrowTypeError("Class extends value is not a constructor or null")
		return Value{Type: VTJSUndefined}
	}

	if superclass.Type == VTJSObject {
		if ctorName := vm.jsObjectStringProperty(superclass, "__js_ctor"); ctorName == "" {
			vm.jsThrowTypeError("Class extends value is not a constructor or null")
			return Value{Type: VTJSUndefined}
		}
	}

	vm.jsSetProto(subclass, superclass)

	derivedProto, deferred := vm.jsMemberGet(subclass, "prototype")
	if deferred || derivedProto.Type != VTJSObject {
		vm.jsThrowTypeError("Class prototype is not available")
		return Value{Type: VTJSUndefined}
	}

	superProto, deferred := vm.jsMemberGet(superclass, "prototype")
	if deferred {
		vm.jsThrowTypeError("Class extends value is not a constructor or null")
		return Value{Type: VTJSUndefined}
	}
	if superProto.Type != VTJSObject && superProto.Type != VTNull {
		vm.jsThrowTypeError("Superclass prototype is not an object or null")
		return Value{Type: VTJSUndefined}
	}

	vm.jsSetProto(derivedProto, superProto)
	return subclass
}

func (vm *VM) jsDefineProperty(target Value, name string, kind int, val Value) {
	if target.Type != VTJSObject && target.Type != VTJSFunction {
		return
	}
	id := target.Num
	props := vm.jsEnsurePropertyMap(id)
	desc, ok := props[name]
	if !ok {
		desc = jsPropertyDescriptor{
			Enumerable:   false,
			Configurable: true,
			Writable:     true,
		}
	}

	switch kind {
	case jsPropertyKindMethod:
		desc.Value = val
		desc.HasValue = true
		desc.Writable = true
	case jsPropertyKindGet:
		desc.Getter = val
		desc.HasGetter = true
		desc.HasValue = false
		desc.Writable = false
	case jsPropertyKindSet:
		desc.Setter = val
		desc.HasSetter = true
		desc.HasValue = false
		desc.Writable = false
	}

	vm.jsSetDescriptor(id, name, desc)
}

func (vm *VM) jsGetDescriptor(objID int64, key string) (jsPropertyDescriptor, bool) {
	props, ok := vm.jsPropertyItems[objID]
	if ok {
		d, exists := props[key]
		if exists {
			return d, true
		}
	}
	obj, ok := vm.jsObjectItems[objID]
	if !ok {
		return jsPropertyDescriptor{}, false
	}
	v, exists := obj[key]
	if !exists {
		return jsPropertyDescriptor{}, false
	}
	return jsDefaultPropertyDescriptor(v), true
}

func (vm *VM) jsSetDescriptor(objID int64, key string, desc jsPropertyDescriptor) {
	props := vm.jsEnsurePropertyMap(objID)
	props[key] = desc
	if desc.HasValue {
		obj, ok := vm.jsObjectItems[objID]
		if !ok {
			obj = make(map[string]Value, 8)
			vm.jsObjectItems[objID] = obj
		}
		obj[key] = desc.Value
	}
	vm.jsInvalidateObjectIC(objID)
}

func (vm *VM) jsCreatePrototypeObject(owner Value) Value {
	protoID := vm.allocJSID()
	vm.jsObjectItems[protoID] = make(map[string]Value, 2)
	vm.jsPropertyItems[protoID] = make(map[string]jsPropertyDescriptor, 2)
	vm.jsSetDescriptor(protoID, "constructor", jsPropertyDescriptor{
		Value:        owner,
		HasValue:     true,
		Enumerable:   false,
		Configurable: true,
		Writable:     true,
	})
	return Value{Type: VTJSObject, Num: protoID}
}

func jsCtorNeedsPrototype(ctorName string) bool {
	switch ctorName {
	case "Array", "Object", "String", "Date", "RegExp", "Enumerator", "VBArray", "Set", "Map", "Promise",
		"ArrayBuffer", "DataView",
		"Int8Array", "Uint8Array", "Uint8ClampedArray",
		"Int16Array", "Uint16Array",
		"Int32Array", "Uint32Array",
		"Float32Array", "Float64Array",
		"BigInt64Array", "BigUint64Array":
		return true
	default:
		return false
	}
}

func (vm *VM) jsGetPrototypeValue(v Value) Value {
	if v.Type == VTJSObject || v.Type == VTJSFunction {
		if obj, ok := vm.jsObjectItems[v.Num]; ok {
			if proto, exists := obj["__js_proto"]; exists {
				return proto
			}
		}
	}
	if v.Type == VTJSFunction {
		if fn, ok := vm.jsFunctionItems[v.Num]; ok && fn != nil && fn.protoID != 0 {
			// Fallback: this is usually the prototype object for instances,
			// but we prefer __js_proto for actual inheritance chain.
			// Actually, for functions, __js_proto IS the parent constructor.
		}
	}
	return Value{Type: VTJSUndefined}
}

func (vm *VM) jsResolveObjectMember(objID int64, member string, visited map[int64]struct{}) (jsPropertyDescriptor, bool) {
	if _, seen := visited[objID]; seen {
		return jsPropertyDescriptor{}, false
	}
	visited[objID] = struct{}{}
	if desc, ok := vm.jsGetDescriptor(objID, member); ok {
		return desc, true
	}
	obj, ok := vm.jsObjectItems[objID]
	if !ok {
		return jsPropertyDescriptor{}, false
	}
	proto, exists := obj["__js_proto"]
	if !exists || (proto.Type != VTJSObject && proto.Type != VTJSFunction) {
		return jsPropertyDescriptor{}, false
	}
	return vm.jsResolveObjectMember(proto.Num, member, visited)
}

func (vm *VM) jsGetIntrinsicPrototype(name string) Value {
	vm.ensureJSRootEnv()
	ctor := vm.jsGetName(name)
	if ctor.Type == VTJSObject || ctor.Type == VTJSFunction {
		if value, deferred := vm.jsMemberGet(ctor, "prototype"); !deferred {
			return value
		}
	}
	return Value{Type: VTJSUndefined}
}

func (vm *VM) jsToDescriptorBoolean(v Value, fallback bool) bool {
	if v.Type == VTJSUndefined {
		return fallback
	}
	return vm.jsTruthy(v)
}

func (vm *VM) jsObjectOwnEnumerableKeys(objID int64) []string {
	names := vm.jsObjectOwnPropertyNames(Value{Type: VTJSObject, Num: objID})
	if len(names) == 0 {
		return nil
	}
	keys := make([]string, 0, len(names))
	for i := 0; i < len(names); i++ {
		desc, hasDesc := vm.jsGetDescriptor(objID, names[i])
		if hasDesc && !desc.Enumerable {
			continue
		}
		keys = append(keys, names[i])
	}
	return keys
}

func (vm *VM) jsJSONParse(text string) Value {
	var payload any
	if err := json.Unmarshal([]byte(text), &payload); err != nil {
		return Value{Type: VTJSUndefined}
	}
	return vm.jsFromGoJSON(payload)
}

func (vm *VM) jsFromGoJSON(payload any) Value {
	switch v := payload.(type) {
	case nil:
		return NewNull()
	case bool:
		return NewBool(v)
	case float64:
		return NewDouble(v)
	case string:
		return NewString(v)
	case []any:
		values := make([]Value, len(v))
		for i := 0; i < len(v); i++ {
			values[i] = vm.jsFromGoJSON(v[i])
		}
		return ValueFromVBArray(NewVBArrayFromValues(0, values))
	case map[string]any:
		objID := vm.allocJSID()
		obj := make(map[string]Value, len(v))
		for key, item := range v {
			obj[key] = vm.jsFromGoJSON(item)
		}
		vm.jsObjectItems[objID] = obj
		props := vm.jsEnsurePropertyMap(objID)
		for key, item := range obj {
			props[key] = jsDefaultPropertyDescriptor(item)
		}
		return Value{Type: VTJSObject, Num: objID}
	default:
		return Value{Type: VTJSUndefined}
	}
}

// jsJSONStringify converts a value into JSON text honoring ES5 toJSON hooks.
func (vm *VM) jsJSONStringify(v Value) string {
	return vm.jsJSONStringifyValue(v)
}

// jsJSONStringifyValue converts a value into JSON text honoring ES5 toJSON hooks.
func (vm *VM) jsJSONStringifyValue(v Value) string {
	v = vm.jsJSONStringifyApplyToJSON(v)
	switch v.Type {
	case VTJSUndefined:
		return "null"
	case VTNull, VTEmpty:
		return "null"
	case VTBool:
		if v.Num != 0 {
			return "true"
		}
		return "false"
	case VTInteger:
		return strconv.FormatInt(v.Num, 10)
	case VTDouble:
		if math.IsNaN(v.Flt) || math.IsInf(v.Flt, 0) {
			return "null"
		}
		return strconv.FormatFloat(v.Flt, 'f', -1, 64)
	case VTString:
		encoded, _ := json.Marshal(v.Str)
		return string(encoded)
	case VTArray:
		if v.Arr == nil {
			return "[]"
		}
		var b strings.Builder
		b.WriteByte('[')
		for i := 0; i < len(v.Arr.Values); i++ {
			if i > 0 {
				b.WriteByte(',')
			}
			b.WriteString(vm.jsJSONStringifyValue(v.Arr.Values[i]))
		}
		b.WriteByte(']')
		return b.String()
	case VTJSObject:
		obj, ok := vm.jsObjectItems[v.Num]
		if !ok {
			return "{}"
		}
		keys := vm.jsObjectOwnEnumerableKeys(v.Num)
		var b strings.Builder
		b.WriteByte('{')
		for i := 0; i < len(keys); i++ {
			if i > 0 {
				b.WriteByte(',')
			}
			k := keys[i]
			encodedKey, _ := json.Marshal(k)
			b.Write(encodedKey)
			b.WriteByte(':')
			b.WriteString(vm.jsJSONStringifyValue(obj[k]))
		}
		b.WriteByte('}')
		return b.String()
	case VTNativeObject:
		// Serialize native Dictionary objects (e.g. produced by G3JSON.LoadFile/Parse)
		// as JSON objects, recursing into nested dictionaries and arrays.
		if dict, ok := vm.dictionaryItems[v.Num]; ok && dict != nil {
			var b strings.Builder
			b.WriteByte('{')
			for i, k := range dict.keys {
				if i > 0 {
					b.WriteByte(',')
				}
				encodedKey, _ := json.Marshal(k.String())
				b.Write(encodedKey)
				b.WriteByte(':')
				b.WriteString(vm.jsJSONStringifyValue(dict.values[i]))
			}
			b.WriteByte('}')
			return b.String()
		}
		// Non-dictionary native: produce a JSON string (best-effort).
		encoded, _ := json.Marshal(vm.valueToString(v))
		return string(encoded)
	default:
		encoded, _ := json.Marshal(vm.valueToString(v))
		return string(encoded)
	}
}

// jsJSONStringifyApplyToJSON invokes object-level toJSON before serialization.
func (vm *VM) jsJSONStringifyApplyToJSON(v Value) Value {
	if v.Type != VTJSObject {
		return v
	}
	toJSON, deferred := vm.jsMemberGet(v, "toJSON")
	if deferred || toJSON.Type != VTJSFunction {
		return v
	}
	result := vm.jsCall(toJSON, v, nil)
	if toJSON.Type == VTJSFunction && len(vm.jsCallStack) > 0 && result.Type == VTJSUndefined {
		return Value{Type: VTJSUndefined}
	}
	return result
}

func (vm *VM) ensureJSRootEnv() {
	if vm.jsActiveEnvID != 0 {
		return
	}
	rootID := vm.allocJSID()
	bindings := make(map[string]Value, 40)
	bindings["Math"] = vm.jsCreateMathObject()
	bindings["Date"] = vm.jsCreateIntrinsicObject("", "Date")
	bindings["RegExp"] = vm.jsCreateIntrinsicObject("", "RegExp")
	bindings["Enumerator"] = vm.jsCreateIntrinsicObject("", "Enumerator")
	bindings["VBArray"] = vm.jsCreateIntrinsicObject("", "VBArray")
	bindings["String"] = vm.jsCreateIntrinsicObject("", "String")
	bindings["Array"] = vm.jsCreateIntrinsicObject("", "Array")
	bindings["Object"] = vm.jsCreateIntrinsicObject("", "Object")
	bindings["JSON"] = vm.jsCreateIntrinsicObject("", "JSON")
	bindings["Promise"] = vm.jsCreatePromiseObject()
	bindings["Number"] = vm.jsCreateNumberObject()
	bindings["Symbol"] = vm.jsCreateSymbolObject()
	bindings["Set"] = vm.jsCreateIntrinsicObject("", "Set")
	bindings["Map"] = vm.jsCreateIntrinsicObject("", "Map")
	// ES6 Binary Data constructors
	bindings["ArrayBuffer"] = vm.jsCreateIntrinsicObject("", "ArrayBuffer")
	bindings["DataView"] = vm.jsCreateIntrinsicObject("", "DataView")
	bindings["Int8Array"] = vm.jsCreateIntrinsicObject("", "Int8Array")
	bindings["Uint8Array"] = vm.jsCreateIntrinsicObject("", "Uint8Array")
	bindings["Uint8ClampedArray"] = vm.jsCreateIntrinsicObject("", "Uint8ClampedArray")
	bindings["Int16Array"] = vm.jsCreateIntrinsicObject("", "Int16Array")
	bindings["Uint16Array"] = vm.jsCreateIntrinsicObject("", "Uint16Array")
	bindings["Int32Array"] = vm.jsCreateIntrinsicObject("", "Int32Array")
	bindings["Uint32Array"] = vm.jsCreateIntrinsicObject("", "Uint32Array")
	bindings["Float32Array"] = vm.jsCreateIntrinsicObject("", "Float32Array")
	bindings["Float64Array"] = vm.jsCreateIntrinsicObject("", "Float64Array")
	bindings["BigInt64Array"] = vm.jsCreateIntrinsicObject("", "BigInt64Array")
	bindings["BigUint64Array"] = vm.jsCreateIntrinsicObject("", "BigUint64Array")
	bindings["NaN"] = NewDouble(math.NaN())
	bindings["Infinity"] = NewDouble(math.Inf(1))
	bindings["undefined"] = Value{Type: VTJSUndefined}
	bindings["isNaN"] = vm.jsCreateIntrinsicObject("", "isNaN")
	bindings["isFinite"] = vm.jsCreateIntrinsicObject("", "isFinite")
	bindings["parseInt"] = vm.jsCreateIntrinsicObject("", "parseInt")
	bindings["parseFloat"] = vm.jsCreateIntrinsicObject("", "parseFloat")
	if evalIdx, ok := GetBuiltinIndex("Eval"); ok {
		bindings["eval"] = Value{Type: VTBuiltin, Num: int64(evalIdx)}
	}
	vm.jsEnvItems[rootID] = &jsEnvFrame{parentID: 0, bindings: bindings}
	vm.jsActiveEnvID = rootID
	vm.jsThisValue = Value{Type: VTJSUndefined}
	vm.jsPopulatePrototypes(bindings)
}

// jsCreateMathObject allocates the global Math object with immutable constants.
func (vm *VM) jsCreateMathObject() Value {
	objID := vm.allocJSID()
	obj := make(map[string]Value, 10)
	obj["__js_type"] = NewString("Math")
	obj["E"] = NewDouble(math.E)
	obj["PI"] = NewDouble(math.Pi)
	obj["LN2"] = NewDouble(math.Ln2)
	obj["LN10"] = NewDouble(math.Ln10)
	obj["LOG2E"] = NewDouble(math.Log2E)
	obj["LOG10E"] = NewDouble(math.Log10E)
	obj["SQRT2"] = NewDouble(math.Sqrt2)
	obj["SQRT1_2"] = NewDouble(1 / math.Sqrt2)
	vm.jsObjectItems[objID] = obj
	vm.jsPropertyItems[objID] = make(map[string]jsPropertyDescriptor, 10)
	return Value{Type: VTJSObject, Num: objID}
}

// jsCreateNumberObject allocates the global Number constructor object with ES6 static methods.
func (vm *VM) jsCreateNumberObject() Value {
	objID := vm.allocJSID()
	obj := make(map[string]Value, 8)
	obj["__js_type"] = NewString("Number")
	obj["__js_ctor"] = NewString("Number")
	// ES6 static constants
	obj["MAX_SAFE_INTEGER"] = NewInteger(9007199254740991)
	obj["MIN_SAFE_INTEGER"] = NewInteger(-9007199254740991)
	obj["MAX_VALUE"] = NewDouble(math.MaxFloat64)
	obj["MIN_VALUE"] = NewDouble(5e-324)
	obj["POSITIVE_INFINITY"] = NewDouble(math.Inf(1))
	obj["NEGATIVE_INFINITY"] = NewDouble(math.Inf(-1))
	obj["NaN"] = NewDouble(math.NaN())
	obj["EPSILON"] = NewDouble(2.220446049250313e-16)
	vm.jsObjectItems[objID] = obj
	vm.jsPropertyItems[objID] = make(map[string]jsPropertyDescriptor, 10)
	vm.jsSetDescriptor(objID, "MAX_SAFE_INTEGER", jsPropertyDescriptor{Value: obj["MAX_SAFE_INTEGER"], HasValue: true, Enumerable: false, Configurable: false, Writable: false})
	vm.jsSetDescriptor(objID, "MIN_SAFE_INTEGER", jsPropertyDescriptor{Value: obj["MIN_SAFE_INTEGER"], HasValue: true, Enumerable: false, Configurable: false, Writable: false})
	vm.jsSetDescriptor(objID, "EPSILON", jsPropertyDescriptor{Value: obj["EPSILON"], HasValue: true, Enumerable: false, Configurable: false, Writable: false})
	return Value{Type: VTJSObject, Num: objID}
}

// jsCreateSymbolObject allocates the global Symbol constructor object and attaches
// well-known symbol properties (Symbol.iterator, Symbol.toStringTag, etc.).
func (vm *VM) jsCreateSymbolObject() Value {
	objID := vm.allocJSID()
	obj := make(map[string]Value, 10)
	obj["__js_type"] = NewString("Symbol")
	obj["__js_ctor"] = NewString("Symbol")
	// Well-known symbols as non-writable, non-enumerable, non-configurable properties
	obj["iterator"] = jsWellKnownSymbolValue(jsWellKnownSymbolIterator, "Symbol.iterator")
	obj["toStringTag"] = jsWellKnownSymbolValue(jsWellKnownSymbolToStringTag, "Symbol.toStringTag")
	obj["species"] = jsWellKnownSymbolValue(jsWellKnownSymbolSpecies, "Symbol.species")
	obj["hasInstance"] = jsWellKnownSymbolValue(jsWellKnownSymbolHasInstance, "Symbol.hasInstance")
	obj["toPrimitive"] = jsWellKnownSymbolValue(jsWellKnownSymbolToPrimitive, "Symbol.toPrimitive")
	vm.jsObjectItems[objID] = obj
	props := make(map[string]jsPropertyDescriptor, 6)
	// Make well-known symbols read-only, non-enumerable, non-configurable
	for _, name := range []string{"iterator", "toStringTag", "species", "hasInstance", "toPrimitive"} {
		props[name] = jsPropertyDescriptor{
			Value: obj[name], HasValue: true,
			Enumerable: false, Configurable: false, Writable: false,
		}
	}
	vm.jsPropertyItems[objID] = props
	return Value{Type: VTJSObject, Num: objID}
}

func (vm *VM) jsCreatePromiseObject() Value {
	objID := vm.allocJSID()
	obj := make(map[string]Value, 4)
	obj["__js_type"] = NewString("Promise")
	obj["__js_ctor"] = NewString("Promise")
	vm.jsObjectItems[objID] = obj
	vm.jsPropertyItems[objID] = make(map[string]jsPropertyDescriptor, 6)

	// Promise.prototype
	proto := vm.jsCreatePrototypeObject(Value{Type: VTJSObject, Num: objID})
	vm.jsSetDescriptor(objID, "prototype", jsPropertyDescriptor{
		Value: proto, HasValue: true, Enumerable: false, Configurable: false, Writable: false,
	})

	// Static methods: Promise.resolve, Promise.reject, Promise.all, Promise.race
	for _, name := range []string{"resolve", "reject", "all", "race"} {
		vm.jsSetDescriptor(objID, name, jsPropertyDescriptor{
			Value:        vm.jsCreateIntrinsicFunction("Promise."+name, "PromiseStatic"+strings.Title(name)),
			HasValue:     true,
			Enumerable:   false,
			Configurable: true,
			Writable:     true,
		})
	}

	return Value{Type: VTJSObject, Num: objID}
}

func (vm *VM) jsCreateIntrinsicFunction(name string, ctorName string) Value {
	objID := vm.allocJSID()
	obj := make(map[string]Value, 3)
	obj["__js_type"] = NewString("Function")
	obj["__js_ctor"] = NewString(ctorName)
	obj["name"] = NewString(name)
	vm.jsObjectItems[objID] = obj
	return Value{Type: VTJSFunction, Num: objID}
}

func (vm *VM) jsCreateIntrinsicObject(typeName string, ctorName string) Value {
	objID := vm.allocJSID()
	obj := make(map[string]Value, 2)
	if typeName != "" {
		obj["__js_type"] = NewString(typeName)
	}
	if ctorName != "" {
		obj["__js_ctor"] = NewString(ctorName)
	}
	vm.jsObjectItems[objID] = obj
	vm.jsPropertyItems[objID] = make(map[string]jsPropertyDescriptor, 4)
	if jsCtorNeedsPrototype(ctorName) {
		ctorVal := Value{Type: VTJSObject, Num: objID}
		proto := vm.jsCreatePrototypeObject(ctorVal)
		vm.jsSetDescriptor(objID, "prototype", jsPropertyDescriptor{
			Value:        proto,
			HasValue:     true,
			Enumerable:   false,
			Configurable: false,
			Writable:     false,
		})
	}
	return Value{Type: VTJSObject, Num: objID}
}

func (vm *VM) jsObjectStringProperty(obj Value, key string) string {
	if obj.Type != VTJSObject && obj.Type != VTJSFunction {
		return ""
	}
	items, ok := vm.jsObjectItems[obj.Num]
	if !ok {
		return ""
	}
	v, ok := items[key]
	if !ok || v.Type != VTString {
		return ""
	}
	return v.Str
}

func (vm *VM) allocJSID() int64 {
	id := vm.nextDynamicNativeID
	vm.nextDynamicNativeID++
	return id
}

func (vm *VM) jsAllocSymbolID() int64 {
	id := vm.jsNextSymbolID
	vm.jsNextSymbolID++
	if vm.jsNextSymbolID <= 0 {
		vm.jsNextSymbolID = 1
	}
	return id
}

func (vm *VM) jsValueMapKey(v Value) string {
	v = resolveCallable(vm, v)
	switch v.Type {
	case VTJSUndefined:
		return "u"
	case VTNull, VTEmpty:
		return "n"
	case VTBool:
		return "b:" + strconv.FormatInt(v.Num, 10)
	case VTInteger:
		f := float64(v.Num)
		if f == 0 {
			return "num:0"
		}
		return "num:" + strconv.FormatFloat(f, 'g', -1, 64)
	case VTDouble:
		if math.IsNaN(v.Flt) {
			return "num:nan"
		}
		if v.Flt == 0 {
			return "num:0"
		}
		return "num:" + strconv.FormatFloat(v.Flt, 'g', -1, 64)
	case VTString:
		return "s:" + v.Str
	case VTJSBigInt:
		if v.Big == nil {
			return "big:0"
		}
		return "big:" + v.Big.String()
	case VTSymbol:
		return "sym:" + strconv.FormatInt(v.Num, 10)
	case VTArray:
		if v.Arr == nil {
			return "arr:nil"
		}
		return "arr:" + fmt.Sprintf("%p", v.Arr)
	case VTJSObject:
		return "jso:" + strconv.FormatInt(v.Num, 10)
	case VTJSFunction:
		return "jsf:" + strconv.FormatInt(v.Num, 10)
	case VTNativeObject:
		return "nat:" + strconv.FormatInt(v.Num, 10)
	case VTObject:
		return "obj:" + strconv.FormatInt(v.Num, 10)
	case VTDate:
		return "date:" + strconv.FormatInt(v.Num, 10)
	default:
		return "t:" + strconv.FormatInt(int64(v.Type), 10) + ":" + v.String()
	}
}

func (vm *VM) jsPropertyKeyFromValue(v Value) string {
	v = resolveCallable(vm, v)
	if v.Type == VTSymbol {
		return jsSymbolPropertyPrefix + strconv.FormatInt(v.Num, 10)
	}
	return vm.valueToString(v)
}

// jsCleanupCollections clears Set/Map backing stores between top-level runs to
// avoid stale growth in long-lived server processes.
func (vm *VM) jsCleanupCollections() {
	for id, setItems := range vm.jsSetItems {
		clear(setItems)
		delete(vm.jsSetItems, id)
	}
	for id, mapItems := range vm.jsMapItems {
		clear(mapItems)
		delete(vm.jsMapItems, id)
	}
}

func (vm *VM) jsCurrentEnv() *jsEnvFrame {
	vm.ensureJSRootEnv()
	return vm.jsEnvItems[vm.jsActiveEnvID]
}

func (vm *VM) jsDeclareName(name string) {
	env := vm.jsCurrentEnv()
	if env == nil {
		return
	}
	if _, ok := env.bindings[name]; ok {
		return
	}
	env.bindings[name] = Value{Type: VTJSUndefined}
}

// jsGetNameFromEnv reads a variable value from the env frame chain starting at envID.
func (vm *VM) jsGetNameFromEnv(envID int64, name string) Value {
	for id := envID; id != 0; {
		env := vm.jsEnvItems[id]
		if env == nil {
			break
		}
		if val, ok := env.bindings[name]; ok {
			return val
		}
		id = env.parentID
	}
	return Value{Type: VTJSUndefined}
}

// jsSetNameInEnv writes a variable value into the env frame chain starting at envID.
// It searches the chain from envID upward for the first frame that has the binding.
func (vm *VM) jsSetNameInEnv(envID int64, name string, val Value) {
	for id := envID; id != 0; {
		env := vm.jsEnvItems[id]
		if env == nil {
			break
		}
		if _, ok := env.bindings[name]; ok {
			env.bindings[name] = val
			return
		}
		id = env.parentID
	}
	// Not found: declare in the given env frame
	if env, ok := vm.jsEnvItems[envID]; ok {
		env.bindings[name] = val
	}
}

// jsGetBlockScopeValue reads from block scopes (innermost first). Returns the value and
// whether it was found. Raises ReferenceError if the variable is in TDZ (const before init).
func (vm *VM) jsGetBlockScopeValue(name string) (Value, bool) {
	for i := len(vm.jsBlockScopes) - 1; i >= 0; i-- {
		if _, exists := vm.jsBlockScopes[i][name]; exists {
			// Check TDZ for const variables (access before initialization).
			if _, inTDZ := vm.jsBlockScopeTDZ[i][name]; inTDZ {
				vm.jsThrowReferenceError(fmt.Sprintf("Cannot access '%s' before initialization", name))
				return Value{Type: VTJSUndefined}, true
			}
			return vm.jsBlockScopes[i][name], true
		}
	}
	return Value{Type: VTJSUndefined}, false
}

// jsSetBlockScopeValue writes into the innermost block scope that declares the name.
// Returns true if found and set. Raises TypeError for const reassignment.
// Raises ReferenceError if const is in TDZ (accessed before initialization).
func (vm *VM) jsSetBlockScopeValue(name string, val Value) bool {
	for i := len(vm.jsBlockScopes) - 1; i >= 0; i-- {
		if _, exists := vm.jsBlockScopes[i][name]; exists {
			// Check TDZ for const before initialization.
			if _, inTDZ := vm.jsBlockScopeTDZ[i][name]; inTDZ {
				vm.jsThrowReferenceError(fmt.Sprintf("Cannot access '%s' before initialization", name))
				return true
			}
			// Check const immutability: once initialized, const cannot be reassigned.
			if _, isConst := vm.jsBlockScopeConst[i][name]; isConst {
				vm.jsThrowTypeError(fmt.Sprintf("Assignment to constant variable '%s'", name))
				return true
			}
			vm.jsBlockScopes[i][name] = val
			return true
		}
	}
	return false
}

func (vm *VM) jsSetName(name string, val Value) {
	vm.ensureJSRootEnv()
	if vm.jsAssignWithBinding(name, val) {
		return
	}
	// Block scopes take precedence for let/const bindings
	if vm.jsBlockScopeDepth > 0 {
		if vm.jsSetBlockScopeValue(name, val) {
			return
		}
	}
	for envID := vm.jsActiveEnvID; envID != 0; {
		env := vm.jsEnvItems[envID]
		if env == nil {
			break
		}
		if _, ok := env.bindings[name]; ok {
			env.bindings[name] = val
			vm.jsSyncArgumentAliasByParam(envID, name, val)
			return
		}
		envID = env.parentID
	}
	if idx, ok := vm.lookupJSGlobalIndex(name); ok {
		vm.Globals[idx] = val
		return
	}

	// In strict mode, assigning to an undeclared variable is a ReferenceError
	if vm.jsStrictMode {
		vm.jsThrowReferenceError(fmt.Sprintf("%s is not defined", name))
		return
	}

	// Non-strict mode: create variable in root/current environment
	root := vm.jsEnvItems[vm.jsActiveEnvID]
	if root != nil {
		root.bindings[name] = val
		vm.jsSyncArgumentAliasByParam(vm.jsActiveEnvID, name, val)
	}
}

func (vm *VM) jsGetName(name string) Value {
	vm.ensureJSRootEnv()
	if strings.EqualFold(name, "this") {
		if vm.jsThisValue.Type == VTJSUninitialized {
			vm.jsThrowReferenceError("jsGetName: Must call super constructor in derived class before accessing 'this'")
			return Value{Type: VTJSUndefined}
		}
		return vm.jsThisValue
	}
	if strings.EqualFold(name, "eval") {
		if idx, ok := GetBuiltinIndex("Eval"); ok {
			return Value{Type: VTBuiltin, Num: int64(idx)}
		}
	}
	// Block scopes take precedence (innermost let/const declarations)
	if vm.jsBlockScopeDepth > 0 {
		if val, found := vm.jsGetBlockScopeValue(name); found {
			return val
		}
	}
	if value, ok := vm.jsResolveWithBinding(name); ok {
		return value
	}
	for envID := vm.jsActiveEnvID; envID != 0; {
		env := vm.jsEnvItems[envID]
		if env == nil {
			break
		}
		if val, ok := env.bindings[name]; ok {
			return val
		}
		envID = env.parentID
	}
	if idx, ok := vm.lookupJSGlobalIndex(name); ok {
		return vm.Globals[idx]
	}
	return Value{Type: VTJSUndefined}
}

func (vm *VM) jsResolveWithBinding(name string) (Value, bool) {
	if len(vm.withStack) == 0 {
		return Value{Type: VTJSUndefined}, false
	}
	for i := len(vm.withStack) - 1; i >= 0; i-- {
		target := vm.withStack[i]
		if !vm.jsHasProperty(target, name) {
			continue
		}
		value, deferred := vm.jsMemberGet(target, name)
		if deferred {
			return Value{Type: VTJSUndefined}, false
		}
		return value, true
	}
	return Value{Type: VTJSUndefined}, false
}

func (vm *VM) jsAssignWithBinding(name string, val Value) bool {
	if len(vm.withStack) == 0 {
		return false
	}
	for i := len(vm.withStack) - 1; i >= 0; i-- {
		target := vm.withStack[i]
		if !vm.jsHasProperty(target, name) {
			continue
		}
		vm.jsMemberSet(target, name, val)
		return true
	}
	return false
}

func (vm *VM) jsHasProperty(target Value, name string) bool {
	switch target.Type {
	case VTJSObject, VTJSFunction:
		_, ok := vm.jsResolveObjectMember(target.Num, name, make(map[int64]struct{}, 4))
		if ok {
			return true
		}
		if obj, hasObj := vm.jsObjectItems[target.Num]; hasObj {
			_, ok = obj[name]
			return ok
		}
		return false
	case VTArray:
		if strings.EqualFold(name, "length") {
			return true
		}
		if target.Arr == nil {
			return false
		}
		idx, err := strconv.Atoi(name)
		if err != nil {
			return false
		}
		adjusted := idx - target.Arr.Lower
		return adjusted >= 0 && adjusted < len(target.Arr.Values)
	case VTString:
		if strings.EqualFold(name, "length") {
			return true
		}
		idx, err := strconv.Atoi(name)
		if err != nil {
			return false
		}
		runes := []rune(target.Str)
		return idx >= 0 && idx < len(runes)
	default:
		return false
	}
}

func (vm *VM) lookupJSGlobalIndex(name string) (int, bool) {
	lowerName := strings.ToLower(name)
	switch lowerName {
	case "response":
		if len(vm.Globals) > 0 {
			return 0, true
		}
	case "request":
		if len(vm.Globals) > 1 {
			return 1, true
		}
	case "server":
		if len(vm.Globals) > 2 {
			return 2, true
		}
	case "session":
		if len(vm.Globals) > 3 {
			return 3, true
		}
	case "application":
		if len(vm.Globals) > 4 {
			return 4, true
		}
	case "objectcontext":
		if len(vm.Globals) > 5 {
			return 5, true
		}
	case "err":
		if len(vm.Globals) > 6 {
			return 6, true
		}
	case "console":
		if len(vm.Globals) > 7 {
			return 7, true
		}
	}

	for i := 0; i < len(vm.globalNames); i++ {
		if vm.globalNames[i] == name {
			return i, true
		}
	}
	if idx, ok := vm.globalNameIndex[lowerName]; ok {
		return idx, true
	}

	// Some execution paths construct a VM without compiler scope metadata
	// (globalNames/globalNameIndex). In that case, resolve builtins by scanning
	// globals for the matching builtin index instead of relying on fixed slots.
	if builtinIdx, ok := BuiltinIndex[lowerName]; ok {
		for i := 0; i < len(vm.Globals); i++ {
			if vm.Globals[i].Type == VTBuiltin && int(vm.Globals[i].Num) == builtinIdx {
				return i, true
			}
		}
	}

	return 0, false
}

func (vm *VM) jsTruthy(v Value) bool {
	switch v.Type {
	case VTJSUndefined, VTNull, VTEmpty:
		return false
	case VTBool:
		return v.Num != 0
	case VTInteger:
		return v.Num != 0
	case VTDouble:
		if math.IsNaN(v.Flt) {
			return false
		}
		return v.Flt != 0
	case VTString:
		return v.Str != ""
	case VTJSBigInt:
		return v.Big != nil && v.Big.Sign() != 0
	default:
		return true
	}
}

func (vm *VM) jsStrictEquals(a Value, b Value) bool {
	if a.Type != b.Type {
		// In JScript, integers and doubles are both "number" type
		if (a.Type == VTInteger || a.Type == VTDouble) && (b.Type == VTInteger || b.Type == VTDouble) {
			return vm.jsToNumber(a).Flt == vm.jsToNumber(b).Flt
		}
		return false
	}
	switch a.Type {
	case VTJSUndefined, VTNull:
		return true
	case VTBool, VTInteger, VTDate, VTNativeObject, VTJSObject, VTJSFunction, VTSymbol:
		return a.Num == b.Num
	case VTDouble:
		return a.Flt == b.Flt
	case VTString:
		return a.Str == b.Str
	case VTJSBigInt:
		return a.Big.Cmp(b.Big) == 0
	default:
		return a.String() == b.String()
	}
}

func (vm *VM) jsTypeOf(v Value) string {
	switch v.Type {
	case VTJSUndefined:
		return "undefined"
	case VTNull:
		return "object"
	case VTBool:
		return "boolean"
	case VTInteger, VTDouble:
		return "number"
	case VTString:
		return "string"
	case VTSymbol:
		return "symbol"
	case VTJSFunction:
		return "function"
	case VTJSObject, VTNativeObject, VTObject, VTArray:
		return "object"
	default:
		return "undefined"
	}
}

// jsAddIntegersNoOverflow adds two int64 values and reports if the result is representable as int64.
func jsAddIntegersNoOverflow(a int64, b int64) (int64, bool) {
	sum := a + b
	if (b > 0 && sum < a) || (b < 0 && sum > a) {
		return 0, false
	}
	return sum, true
}

// jsSubtractIntegersNoOverflow subtracts two int64 values and reports if the result is representable as int64.
func jsSubtractIntegersNoOverflow(a int64, b int64) (int64, bool) {
	diff := a - b
	if (b < 0 && diff < a) || (b > 0 && diff > a) {
		return 0, false
	}
	return diff, true
}

// jsIncrementNumberValue increments a numeric value preserving VTInteger whenever possible.
func (vm *VM) jsIncrementNumberValue(v Value) Value {
	if v.Type == VTInteger {
		if next, ok := jsAddIntegersNoOverflow(v.Num, 1); ok {
			return NewInteger(next)
		}
		return NewDouble(float64(v.Num) + 1)
	}
	next := vm.jsToNumber(v)
	next.Flt++
	return next
}

// jsDecrementNumberValue decrements a numeric value preserving VTInteger whenever possible.
func (vm *VM) jsDecrementNumberValue(v Value) Value {
	if v.Type == VTInteger {
		if next, ok := jsSubtractIntegersNoOverflow(v.Num, 1); ok {
			return NewInteger(next)
		}
		return NewDouble(float64(v.Num) - 1)
	}
	next := vm.jsToNumber(v)
	next.Flt--
	return next
}

// jsAddValues implements JScript '+' behavior for string concatenation and numeric addition.
func (vm *VM) jsAddValues(a Value, b Value) Value {
	a = resolveCallable(vm, a)
	b = resolveCallable(vm, b)
	if a.Type == VTString || b.Type == VTString {
		sa := vm.jsConcatString(a)
		sb := vm.jsConcatString(b)
		total := len(sa) + len(sb)
		if !vm.jsEnsureStringSize(total) || !vm.jsChargeStringWork(total) {
			return Value{Type: VTJSUndefined}
		}
		return NewString(sa + sb)
	}
	if a.Type == VTInteger && b.Type == VTInteger {
		if sum, ok := jsAddIntegersNoOverflow(a.Num, b.Num); ok {
			return NewInteger(sum)
		}
		return NewDouble(float64(a.Num) + float64(b.Num))
	}
	return NewDouble(vm.jsToNumber(a).Flt + vm.jsToNumber(b).Flt)
}

// jsConcatString converts one value to the JScript string form used by '+' concatenation.
func (vm *VM) jsConcatString(v Value) string {
	v = resolveCallable(vm, v)
	if arr, ok := vm.jsAsConcatArray(v); ok {
		return vm.jsArrayToString(arr)
	}
	return vm.valueToString(v)
}

// jsAsConcatArray resolves supported array-like bridge values to a VTArray source.
func (vm *VM) jsAsConcatArray(v Value) (Value, bool) {
	if v.Type == VTArray && v.Arr != nil {
		return v, true
	}
	if v.Type != VTJSObject {
		return Value{Type: VTJSUndefined}, false
	}
	items, ok := vm.jsObjectItems[v.Num]
	if !ok {
		return Value{Type: VTJSUndefined}, false
	}
	vbSource, ok := items["__js_vbarray_source"]
	if !ok {
		return Value{Type: VTJSUndefined}, false
	}
	converted := vm.jsVBArrayToJSArray(vbSource)
	if converted.Type != VTArray || converted.Arr == nil {
		return Value{Type: VTJSUndefined}, false
	}
	return converted, true
}

// jsArrayToString returns the ES3/ES5 Array toString()-style comma-joined representation.
func (vm *VM) jsArrayToString(v Value) string {
	if v.Type != VTArray || v.Arr == nil || len(v.Arr.Values) == 0 {
		return ""
	}
	parts := make([]string, len(v.Arr.Values))
	totalSize := 0
	for i := 0; i < len(v.Arr.Values); i++ {
		item := v.Arr.Values[i]
		if item.Type == VTJSUndefined || item.Type == VTNull {
			parts[i] = ""
		} else {
			parts[i] = vm.jsConcatString(item)
		}
		totalSize += len(parts[i])
		if i > 0 {
			totalSize++
		}
		if !vm.jsEnsureStringSize(totalSize) {
			return ""
		}
	}
	if !vm.jsChargeStringWork(totalSize) {
		return ""
	}
	return strings.Join(parts, ",")
}

// jsEnsureStringSize guards JScript string-producing operations against runaway growth.
func (vm *VM) jsEnsureStringSize(size int) bool {
	if size <= jsMaxStringBytes {
		return true
	}
	vm.raise(vbscript.OutOfStringSpace, fmt.Sprintf("JScript string size exceeded %d bytes", jsMaxStringBytes))
	return false
}

// jsChargeStringWork tracks cumulative JScript string output work per Run() to stop pathological growth loops.
func (vm *VM) jsChargeStringWork(size int) bool {
	if size <= 0 {
		return true
	}
	vm.jsStringWorkBytes += int64(size)
	if vm.jsStringWorkBytes <= jsMaxStringWorkBytes {
		return true
	}
	vm.raise(vbscript.OutOfStringSpace, fmt.Sprintf("JScript cumulative string work exceeded %d bytes", jsMaxStringWorkBytes))
	return false
}

func (vm *VM) jsCreateClosure(template Value) Value {
	if template.Type != VTJSFunctionTemplate && template.Type != VTJSArrowFunctionTemplate {
		return Value{Type: VTJSUndefined}
	}
	id := vm.allocJSID()
	fnVal := Value{Type: VTJSFunction, Num: id}
	proto := vm.jsCreatePrototypeObject(fnVal)
	params := make([]string, 0, len(template.Names))
	restParam := ""
	localCount := 0
	isClassConstructor := false
	isStrict := false
	isDerived := false
	isGenerator := false
	isAsync := false
	for i := 0; i < len(template.Names); i++ {
		name := template.Names[i]
		if strings.HasPrefix(name, "__js_local_count__:") {
			if n, err := strconv.Atoi(strings.TrimPrefix(name, "__js_local_count__:")); err == nil && n > 0 {
				localCount = n
			}
			continue
		}
		if strings.HasPrefix(name, jsRestParamPrefix) {
			restParam = strings.TrimPrefix(name, jsRestParamPrefix)
			continue
		}
		if name == jsClassConstructorFlag {
			isClassConstructor = true
			continue
		}
		if name == jsStrictModeFlag {
			isStrict = true
			continue
		}
		if name == jsGeneratorFlag {
			isGenerator = true
			continue
		}
		if name == jsAsyncFlag {
			isAsync = true
			continue
		}
		if name == jsDerivedConstructorFlag {
			isDerived = true
			continue
		}
		params = append(params, name)
	}
	fnObj := &jsFunctionObject{
		name:               template.Str,
		params:             params,
		restParam:          restParam,
		localCount:         localCount,
		startIP:            int(template.Num),
		endIP:              int(template.Flt),
		envID:              vm.jsActiveEnvID,
		protoID:            proto.Num,
		isClassConstructor: isClassConstructor,
		isStrict:           isStrict,
		isDerived:          isDerived,
		isAsync:            isAsync,
		isGenerator:        isGenerator,
	}
	// Arrow functions capture the current 'this' value lexically at creation time.
	if template.Type == VTJSArrowFunctionTemplate {
		fnObj.isArrow = true
		fnObj.capturedThis = vm.jsThisValue
	}
	vm.jsFunctionItems[id] = fnObj
	vm.jsObjectItems[id] = make(map[string]Value, 2)
	vm.jsPropertyItems[id] = make(map[string]jsPropertyDescriptor, 2)
	vm.jsSetDescriptor(id, "prototype", jsPropertyDescriptor{
		Value:        proto,
		HasValue:     true,
		Enumerable:   false,
		Configurable: false,
		Writable:     false,
	})
	return fnVal
}

func (vm *VM) jsPrepareLocalFrame(localCount int, savedSP int) bool {
	if localCount <= 0 {
		vm.fp = savedSP + 1
		vm.sp = savedSP
		return true
	}
	base := savedSP + 1
	end := base + localCount - 1
	if end >= len(vm.stack) {
		vm.raise(vbscript.StackOverflow, "JScript local frame exceeds VM stack capacity")
		return false
	}
	for i := base; i <= end; i++ {
		vm.stack[i] = Value{Type: VTJSUndefined}
	}
	vm.fp = base
	vm.sp = end
	return true
}

func (vm *VM) jsBeginFunctionCall(fn Value, thisVal Value, args []Value, ctorObj Value, isCtor bool, newTarget Value, isSuperCall bool) bool {
	// Debug:
	// fmt.Printf("jsBeginFunctionCall: fn=%v, isCtor=%v, isSuper=%v\n", fn, isCtor, isSuperCall)
	closure, ok := vm.jsFunctionItems[fn.Num]
	if !ok || closure == nil {
		return false
	}
	// Arrow functions ignore the caller-supplied 'this' and use the lexically captured one.
	if closure.isArrow {
		thisVal = closure.capturedThis
	}
	frame := jsCallFrame{
		returnIP:     vm.ip,
		envID:        vm.jsActiveEnvID,
		savedFP:      vm.fp,
		fn:           fn,
		thisVal:      vm.jsThisValue,
		newTarget:    vm.jsNewTarget,
		tryDepth:     len(vm.jsTryStack),
		savedSP:      vm.sp,
		isCtor:       isCtor,
		ctorObj:      ctorObj,
		jsStrictMode: vm.jsStrictMode,
		isSuperCall:  isSuperCall,
	}
	vm.jsCallStack = append(vm.jsCallStack, frame)
	vm.jsStrictMode = closure.isStrict
	vm.jsNewTarget = newTarget
	envID := vm.allocJSID()
	bindings := make(map[string]Value, len(closure.params)+2)
	for i := 0; i < len(closure.params); i++ {
		if i < len(args) {
			bindings[closure.params[i]] = args[i]
		} else {
			bindings[closure.params[i]] = Value{Type: VTJSUndefined}
		}
	}
	if closure.restParam != "" {
		restValues := make([]Value, 0)
		if len(args) > len(closure.params) {
			restValues = append(restValues, args[len(closure.params):]...)
		}
		bindings[closure.restParam] = ValueFromVBArray(NewVBArrayFromValues(0, restValues))
	}
	if _, hasArguments := bindings["arguments"]; !hasArguments {
		argumentsObject := vm.jsCreateArgumentsObject(args, closure.params, envID)
		bindings["arguments"] = argumentsObject
	}
	vm.jsEnvItems[envID] = &jsEnvFrame{parentID: closure.envID, bindings: bindings}
	vm.jsActiveEnvID = envID
	vm.jsThisValue = thisVal
	if !vm.jsPrepareLocalFrame(closure.localCount, frame.savedSP) {
		vm.jsCallStack = vm.jsCallStack[:len(vm.jsCallStack)-1]
		vm.jsActiveEnvID = frame.envID
		vm.jsThisValue = frame.thisVal
		vm.jsNewTarget = frame.newTarget
		vm.jsStrictMode = frame.jsStrictMode
		vm.fp = frame.savedFP
		vm.sp = frame.savedSP
		return false
	}
	vm.ip = closure.startIP
	return true
}

// jsEnvHasCapturedClosures reports whether any closure currently captures envID.
func (vm *VM) jsEnvHasCapturedClosures(envID int64) bool {
	if envID == 0 {
		return false
	}
	for _, fn := range vm.jsFunctionItems {
		if fn != nil && fn.envID == envID {
			return true
		}
	}
	return false
}

// jsRefreshArgumentsObject rewrites one existing arguments object in place.
func (vm *VM) jsRefreshArgumentsObject(objID int64, args []Value, params []string, envID int64) {
	obj, ok := vm.jsObjectItems[objID]
	if !ok {
		obj = make(map[string]Value, len(args)+1)
		vm.jsObjectItems[objID] = obj
	} else {
		clear(obj)
	}
	for i := 0; i < len(args); i++ {
		obj[strconv.Itoa(i)] = args[i]
	}
	obj["length"] = NewInteger(int64(len(args)))

	if len(params) > 0 && len(args) > 0 {
		alias, ok := vm.jsArgumentsItems[objID]
		if !ok || alias == nil {
			alias = &jsArgumentsBinding{}
			vm.jsArgumentsItems[objID] = alias
		}
		alias.envID = envID
		if alias.indexToParam == nil {
			alias.indexToParam = make(map[string]string, len(params))
		} else {
			clear(alias.indexToParam)
		}
		if alias.paramToIndex == nil {
			alias.paramToIndex = make(map[string]string, len(params))
		} else {
			clear(alias.paramToIndex)
		}
		max := len(params)
		if len(args) < max {
			max = len(args)
		}
		for i := 0; i < max; i++ {
			key := strconv.Itoa(i)
			paramName := params[i]
			alias.indexToParam[key] = paramName
			if _, exists := alias.paramToIndex[paramName]; !exists {
				alias.paramToIndex[paramName] = key
			}
		}
	} else {
		delete(vm.jsArgumentsItems, objID)
	}
}

// jsTailCallValue replaces the current JScript call frame with one new call target.
func (vm *VM) jsTailCallValue(callee Value, thisVal Value, args []Value) bool {
	if callee.Type != VTJSFunction || len(vm.jsCallStack) == 0 {
		return false
	}

	closure, ok := vm.jsFunctionItems[callee.Num]
	if !ok || closure == nil || closure.isBound {
		return false
	}

	if closure.isArrow {
		thisVal = closure.capturedThis
	}

	frame := &vm.jsCallStack[len(vm.jsCallStack)-1]

	canReuseEnv := vm.jsActiveEnvID != 0 && !vm.jsEnvHasCapturedClosures(vm.jsActiveEnvID)
	envID := vm.jsActiveEnvID
	bindings := make(map[string]Value, len(closure.params)+2)
	reusedArgs := Value{Type: VTJSUndefined}
	if canReuseEnv {
		if env, ok := vm.jsEnvItems[envID]; ok && env != nil {
			if env.bindings == nil {
				env.bindings = make(map[string]Value, len(closure.params)+2)
			}
			bindings = env.bindings
			if existingArgs, exists := bindings["arguments"]; exists {
				reusedArgs = existingArgs
			}
			clear(bindings)
			env.parentID = closure.envID
		} else {
			canReuseEnv = false
		}
	}
	if !canReuseEnv {
		envID = vm.allocJSID()
		vm.jsEnvItems[envID] = &jsEnvFrame{parentID: closure.envID, bindings: bindings}
	}

	for i := 0; i < len(closure.params); i++ {
		if i < len(args) {
			bindings[closure.params[i]] = args[i]
		} else {
			bindings[closure.params[i]] = Value{Type: VTJSUndefined}
		}
	}
	if closure.restParam != "" {
		restValues := make([]Value, 0)
		if len(args) > len(closure.params) {
			restValues = append(restValues, args[len(closure.params):]...)
		}
		bindings[closure.restParam] = ValueFromVBArray(NewVBArrayFromValues(0, restValues))
	}
	if _, hasArguments := bindings["arguments"]; !hasArguments {
		if canReuseEnv {
			if reusedArgs.Type == VTJSObject {
				vm.jsRefreshArgumentsObject(reusedArgs.Num, args, closure.params, envID)
				bindings["arguments"] = reusedArgs
			} else {
				bindings["arguments"] = vm.jsCreateArgumentsObject(args, closure.params, envID)
			}
		} else {
			bindings["arguments"] = vm.jsCreateArgumentsObject(args, closure.params, envID)
		}
	}

	if len(vm.jsTryStack) > frame.tryDepth {
		vm.jsTryStack = vm.jsTryStack[:frame.tryDepth]
	}

	vm.jsActiveEnvID = envID
	vm.jsThisValue = thisVal
	vm.ip = closure.startIP
	if !vm.jsPrepareLocalFrame(closure.localCount, frame.savedSP) {
		vm.jsActiveEnvID = frame.envID
		vm.jsThisValue = frame.thisVal
		vm.jsNewTarget = frame.newTarget
		vm.jsStrictMode = frame.jsStrictMode
		vm.fp = frame.savedFP
		vm.sp = frame.savedSP
		vm.ip = frame.returnIP
		return false
	}
	return true
}

func (vm *VM) jsCreateArgumentsObject(args []Value, params []string, envID int64) Value {
	objID := vm.allocJSID()
	obj := make(map[string]Value, len(args)+1)
	for i := 0; i < len(args); i++ {
		obj[strconv.Itoa(i)] = args[i]
	}
	obj["length"] = NewInteger(int64(len(args)))
	vm.jsObjectItems[objID] = obj
	vm.jsPropertyItems[objID] = make(map[string]jsPropertyDescriptor, len(args)+1)
	if len(params) > 0 && len(args) > 0 {
		alias := &jsArgumentsBinding{
			envID:        envID,
			indexToParam: make(map[string]string, len(params)),
			paramToIndex: make(map[string]string, len(params)),
		}
		max := len(params)
		if len(args) < max {
			max = len(args)
		}
		for i := 0; i < max; i++ {
			key := strconv.Itoa(i)
			paramName := params[i]
			alias.indexToParam[key] = paramName
			if _, exists := alias.paramToIndex[paramName]; !exists {
				alias.paramToIndex[paramName] = key
			}
		}
		if vm.jsArgumentsItems == nil {
			vm.jsArgumentsItems = make(map[int64]*jsArgumentsBinding, 8)
		}
		vm.jsArgumentsItems[objID] = alias
	}
	return Value{Type: VTJSObject, Num: objID}
}

func (vm *VM) jsSyncArgumentAliasByParam(envID int64, name string, value Value) {
	if len(vm.jsArgumentsItems) == 0 {
		return
	}
	for objID, alias := range vm.jsArgumentsItems {
		if alias == nil || alias.envID != envID {
			continue
		}
		idxStr, ok := alias.paramToIndex[name]
		if !ok {
			continue
		}
		if obj, exists := vm.jsObjectItems[objID]; exists {
			obj[idxStr] = value
		}
	}
}
func (vm *VM) jsSetAliasedArgumentValue(objID int64, key string, value Value) bool {
	alias, ok := vm.jsArgumentsItems[objID]
	if !ok || alias == nil {
		return false
	}
	paramName, ok := alias.indexToParam[key]
	if !ok {
		return false
	}
	env := vm.jsEnvItems[alias.envID]
	if env == nil {
		return false
	}
	env.bindings[paramName] = value
	if obj, exists := vm.jsObjectItems[objID]; exists {
		obj[key] = value
	}
	return true
}

// jsGetAliasedArgumentValue reads a live argument value from the env frame
// for an arguments binding, so that mutations of named params are visible via arguments[n].
func (vm *VM) jsGetAliasedArgumentValue(objID int64, key string) (Value, bool) {
	alias, ok := vm.jsArgumentsItems[objID]
	if !ok || alias == nil {
		return Value{Type: VTJSUndefined}, false
	}
	paramName, ok := alias.indexToParam[key]
	if !ok {
		return Value{Type: VTJSUndefined}, false
	}
	env := vm.jsEnvItems[alias.envID]
	if env == nil {
		return Value{Type: VTJSUndefined}, false
	}
	val, exists := env.bindings[paramName]
	if !exists {
		return Value{Type: VTJSUndefined}, false
	}
	return val, true
}

func (vm *VM) jsExtractApplyArgs(argArray Value) []Value {
	switch argArray.Type {
	case VTJSUndefined, VTNull:
		return nil
	case VTArray:
		if argArray.Arr == nil || len(argArray.Arr.Values) == 0 {
			return nil
		}
		return argArray.Arr.Values
	case VTJSObject:
		obj, ok := vm.jsObjectItems[argArray.Num]
		if !ok || len(obj) == 0 {
			return nil
		}
		lengthVal, hasLength := obj["length"]
		if !hasLength {
			return nil
		}
		lengthNum := int(vm.jsToNumber(lengthVal).Flt)
		if lengthNum <= 0 {
			return nil
		}
		out := make([]Value, lengthNum)
		for i := 0; i < lengthNum; i++ {
			key := strconv.Itoa(i)
			if v, exists := obj[key]; exists {
				out[i] = v
			} else {
				out[i] = Value{Type: VTJSUndefined}
			}
		}
		return out
	default:
		return nil
	}
}

func (vm *VM) jsArrayFlat(vals []Value, depth int) []Value {
	if depth < 1 {
		return append([]Value(nil), vals...)
	}
	out := make([]Value, 0, len(vals))
	for _, v := range vals {
		if v.Type == VTArray && v.Arr != nil {
			flattened := vm.jsArrayFlat(v.Arr.Values, depth-1)
			out = append(out, flattened...)
		} else {
			out = append(out, v)
		}
	}
	return out
}

func (vm *VM) jsBindFunction(fn Value, thisArg Value, boundArgs []Value) Value {
	if fn.Type != VTJSFunction {
		return Value{Type: VTJSUndefined}
	}
	id := vm.allocJSID()
	bound := &jsFunctionObject{
		name:      "bound",
		isBound:   true,
		boundFn:   fn,
		boundThis: thisArg,
	}
	if len(boundArgs) > 0 {
		bound.boundArgs = append([]Value(nil), boundArgs...)
	}
	proto := vm.jsCreatePrototypeObject(Value{Type: VTJSFunction, Num: id})
	bound.protoID = proto.Num
	vm.jsFunctionItems[id] = bound
	vm.jsObjectItems[id] = make(map[string]Value, 2)
	vm.jsPropertyItems[id] = make(map[string]jsPropertyDescriptor, 2)
	vm.jsSetDescriptor(id, "prototype", jsPropertyDescriptor{
		Value:        proto,
		HasValue:     true,
		Enumerable:   false,
		Configurable: false,
		Writable:     false,
	})
	return Value{Type: VTJSFunction, Num: id}
}

func jsClampIndex(idx int, length int) int {
	if idx < 0 {
		return 0
	}
	if idx > length {
		return length
	}
	return idx
}

func jsNormalizeRelativeIndex(idx int, length int) int {
	if idx < 0 {
		idx = length + idx
	}
	if idx < 0 {
		return 0
	}
	if idx > length {
		return length
	}
	return idx
}

func (vm *VM) jsParseIntES5(args []Value) Value {
	if len(args) == 0 {
		return NewDouble(math.NaN())
	}
	s := strings.TrimSpace(vm.valueToString(args[0]))
	if s == "" {
		return NewDouble(math.NaN())
	}
	sign := 1.0
	if s[0] == '+' || s[0] == '-' {
		if s[0] == '-' {
			sign = -1.0
		}
		s = s[1:]
	}
	if s == "" {
		return NewDouble(math.NaN())
	}
	radix := 0
	if len(args) > 1 {
		r := int(vm.jsToNumber(args[1]).Flt)
		if r != 0 {
			if r < 2 || r > 36 {
				return NewDouble(math.NaN())
			}
			radix = r
		}
	}
	if radix == 0 {
		if len(s) >= 2 && (s[0:2] == "0x" || s[0:2] == "0X") {
			radix = 16
			s = s[2:]
		} else {
			radix = 10
		}
	} else if radix == 16 && len(s) >= 2 && (s[0:2] == "0x" || s[0:2] == "0X") {
		s = s[2:]
	}
	if s == "" {
		return NewDouble(math.NaN())
	}
	value := 0.0
	digits := 0
	for i := 0; i < len(s); i++ {
		ch := s[i]
		d := -1
		switch {
		case ch >= '0' && ch <= '9':
			d = int(ch - '0')
		case ch >= 'a' && ch <= 'z':
			d = int(ch-'a') + 10
		case ch >= 'A' && ch <= 'Z':
			d = int(ch-'A') + 10
		}
		if d < 0 || d >= radix {
			break
		}
		value = value*float64(radix) + float64(d)
		digits++
	}
	if digits == 0 {
		return NewDouble(math.NaN())
	}
	return NewDouble(sign * value)
}

// jsFunctionExpectedLength returns the ES5 Function.length value.
func (vm *VM) jsFunctionExpectedLength(fn Value) int {
	if fn.Type != VTJSFunction {
		return 0
	}
	closure, ok := vm.jsFunctionItems[fn.Num]
	if !ok || closure == nil {
		return 0
	}
	if closure.isBound {
		base := vm.jsFunctionExpectedLength(closure.boundFn)
		remaining := base - len(closure.boundArgs)
		if remaining < 0 {
			return 0
		}
		return remaining
	}
	return len(closure.params)
}

// jsObjectToStringTag returns the canonical Object.prototype.toString tag.
func (vm *VM) jsObjectToStringTag(v Value) string {
	switch v.Type {
	case VTJSUndefined:
		return "[object Undefined]"
	case VTNull:
		return "[object Null]"
	case VTArray:
		return "[object Array]"
	case VTDate:
		return "[object Date]"
	case VTJSFunction:
		return "[object Function]"
	case VTString:
		return "[object String]"
	case VTSymbol:
		return "[object Symbol]"
	case VTBool:
		return "[object Boolean]"
	case VTInteger, VTDouble:
		return "[object Number]"
	case VTJSObject:
		tag := vm.jsObjectStringProperty(v, "__js_type")
		if tag == "" {
			tag = vm.jsObjectStringProperty(v, "__js_ctor")
		}
		switch tag {
		case "Array", "Date", "Function", "RegExp", "Math", "JSON", "Enumerator", "VBArray", "String", "Number", "Boolean", "Object", "Set", "Map",
			"ArrayBuffer", "DataView",
			"Int8Array", "Uint8Array", "Uint8ClampedArray",
			"Int16Array", "Uint16Array",
			"Int32Array", "Uint32Array",
			"Float32Array", "Float64Array",
			"BigInt64Array", "BigUint64Array":
			return "[object " + tag + "]"
		default:
			return "[object Object]"
		}
	default:
		return "[object Object]"
	}
}

// jsArrayLikeLength resolves the effective ES5 array-like length from an object.
func (vm *VM) jsArrayLikeLength(target Value) (int, bool, bool) {
	switch target.Type {
	case VTArray:
		if target.Arr == nil {
			return 0, true, false
		}
		return len(target.Arr.Values), true, false
	case VTJSObject:
		lengthVal, deferred := vm.jsMemberGet(target, "length")
		if deferred {
			return 0, false, true
		}
		if lengthVal.Type == VTJSUndefined {
			return 0, false, false
		}
		n := int(vm.jsToNumber(lengthVal).Flt)
		if n < 0 {
			n = 0
		}
		return n, true, false
	default:
		return 0, false, false
	}
}

// jsArrayLikeHasIndex reports whether an array-like value has an element index.
func (vm *VM) jsArrayLikeHasIndex(target Value, idx int) bool {
	if idx < 0 {
		return false
	}
	if target.Type == VTArray {
		if target.Arr == nil {
			return false
		}
		return idx < len(target.Arr.Values)
	}
	if target.Type != VTJSObject {
		return false
	}
	return vm.jsHasProperty(target, strconv.Itoa(idx))
}

// jsArrayLikeGetIndex resolves one element from an array-like value.
func (vm *VM) jsArrayLikeGetIndex(target Value, idx int) (Value, bool) {
	if idx < 0 {
		return Value{Type: VTJSUndefined}, false
	}
	if target.Type == VTArray {
		if target.Arr == nil || idx >= len(target.Arr.Values) {
			return Value{Type: VTJSUndefined}, false
		}
		return target.Arr.Values[idx], true
	}
	if target.Type != VTJSObject {
		return Value{Type: VTJSUndefined}, false
	}
	v, deferred := vm.jsMemberGet(target, strconv.Itoa(idx))
	if deferred {
		return Value{Type: VTJSUndefined}, false
	}
	if v.Type == VTJSUndefined && !vm.jsArrayLikeHasIndex(target, idx) {
		return Value{Type: VTJSUndefined}, false
	}
	return v, true
}

func (vm *VM) jsParseFloatES5(args []Value) Value {
	if len(args) == 0 {
		return NewDouble(math.NaN())
	}
	s := strings.TrimSpace(vm.valueToString(args[0]))
	if s == "" {
		return NewDouble(math.NaN())
	}
	if strings.HasPrefix(s, "Infinity") || strings.HasPrefix(s, "+Infinity") {
		return NewDouble(math.Inf(1))
	}
	if strings.HasPrefix(s, "-Infinity") {
		return NewDouble(math.Inf(-1))
	}
	i := 0
	if i < len(s) && (s[i] == '+' || s[i] == '-') {
		i++
	}
	intStart := i
	for i < len(s) && s[i] >= '0' && s[i] <= '9' {
		i++
	}
	hasInt := i > intStart
	if i < len(s) && s[i] == '.' {
		i++
		fracStart := i
		for i < len(s) && s[i] >= '0' && s[i] <= '9' {
			i++
		}
		hasInt = hasInt || (i > fracStart)
	}
	if !hasInt {
		return NewDouble(math.NaN())
	}
	end := i
	if i < len(s) && (s[i] == 'e' || s[i] == 'E') {
		expPos := i
		i++
		if i < len(s) && (s[i] == '+' || s[i] == '-') {
			i++
		}
		expStart := i
		for i < len(s) && s[i] >= '0' && s[i] <= '9' {
			i++
		}
		if i > expStart {
			end = i
		} else {
			end = expPos
		}
	}
	parsed, err := strconv.ParseFloat(s[:end], 64)
	if err != nil {
		return NewDouble(math.NaN())
	}
	return NewDouble(parsed)
}

func (vm *VM) jsObjectHasOwnProperty(target Value, key string) bool {
	switch target.Type {
	case VTJSObject, VTJSFunction:
		if _, ok := vm.jsGetDescriptor(target.Num, key); ok {
			return true
		}
		obj, ok := vm.jsObjectItems[target.Num]
		if !ok {
			return false
		}
		_, exists := obj[key]
		return exists
	case VTArray:
		if strings.EqualFold(key, "length") {
			return true
		}
		if target.Arr == nil {
			return false
		}
		i, err := strconv.Atoi(key)
		if err != nil {
			return false
		}
		idx := i - target.Arr.Lower
		return idx >= 0 && idx < len(target.Arr.Values)
	case VTString:
		if strings.EqualFold(key, "length") {
			return true
		}
		i, err := strconv.Atoi(key)
		if err != nil {
			return false
		}
		r := []rune(target.Str)
		return i >= 0 && i < len(r)
	default:
		return false
	}
}

func (vm *VM) jsObjectPropertyIsEnumerable(target Value, key string) bool {
	switch target.Type {
	case VTJSObject, VTJSFunction:
		d, ok := vm.jsGetDescriptor(target.Num, key)
		if !ok {
			return false
		}
		return d.Enumerable
	case VTArray, VTString:
		return vm.jsObjectHasOwnProperty(target, key)
	default:
		return false
	}
}

func (vm *VM) jsObjectIsPrototypeOf(proto Value, candidate Value) bool {
	if proto.Type != VTJSObject && proto.Type != VTJSFunction {
		return false
	}
	seen := make(map[int64]struct{}, 4)
	current := vm.jsGetPrototypeValue(candidate)
	for current.Type == VTJSObject {
		if current.Num == proto.Num {
			return true
		}
		if _, ok := seen[current.Num]; ok {
			break
		}
		seen[current.Num] = struct{}{}
		obj, ok := vm.jsObjectItems[current.Num]
		if !ok {
			break
		}
		next, ok := obj["__js_proto"]
		if !ok {
			break
		}
		current = next
	}
	return false
}

func (vm *VM) jsReturn(retVal Value) {
	if len(vm.jsCallStack) == 0 {
		vm.push(retVal)
		return
	}
	frame := vm.jsCallStack[len(vm.jsCallStack)-1]
	vm.jsCallStack = vm.jsCallStack[:len(vm.jsCallStack)-1]
	if len(vm.jsTryStack) > frame.tryDepth {
		vm.jsTryStack = vm.jsTryStack[:frame.tryDepth]
	}
	vm.jsActiveEnvID = frame.envID
	vm.jsThisValue = frame.thisVal
	vm.jsNewTarget = frame.newTarget
	vm.jsStrictMode = frame.jsStrictMode
	vm.ip = frame.returnIP
	vm.fp = frame.savedFP
	vm.sp = frame.savedSP
	if frame.isSuperCall {
		// This was a super() call in a constructor. Assign the result to 'this'.
		newThis := retVal
		if retVal.Type != VTJSObject && retVal.Type != VTJSFunction && retVal.Type != VTObject && retVal.Type != VTNativeObject && retVal.Type != VTArray {
			newThis = frame.ctorObj
		}
		vm.jsThisValue = newThis
		if len(vm.jsCallStack) > 0 {
			vm.jsCallStack[len(vm.jsCallStack)-1].thisVal = newThis
			if vm.jsCallStack[len(vm.jsCallStack)-1].ctorObj.Type == VTJSUninitialized {
				vm.jsCallStack[len(vm.jsCallStack)-1].ctorObj = newThis
			}
		}
		vm.push(newThis)
		return
	}

	if frame.isCtor {
		switch retVal.Type {
		case VTJSObject, VTJSFunction, VTObject, VTNativeObject, VTArray:
			vm.push(retVal)
			return
		default:
			vm.push(frame.ctorObj)
			return
		}
	}
	vm.push(retVal)
}

func (vm *VM) jsMemberGet(target Value, member string) (Value, bool) {
	if target.Type == VTJSUninitialized {
		vm.jsThrowReferenceError("Must call super constructor in derived class before accessing 'this'")
		return Value{Type: VTJSUndefined}, false
	}
	switch target.Type {
	case VTNativeObject:
		return vm.dispatchMemberGet(target, member), false
	case VTString:
		if strings.EqualFold(member, "length") {
			return NewInteger(jsStringLength(target.Str)), false
		}
		proto := vm.jsGetIntrinsicPrototype("String")
		if proto.Type == VTJSObject {
			desc, hasDesc := vm.jsResolveObjectMember(proto.Num, member, make(map[int64]struct{}, 4))
			if hasDesc {
				if desc.HasGetter {
					result := vm.jsCall(desc.Getter, target, nil)
					if desc.Getter.Type == VTJSFunction && len(vm.jsCallStack) > 0 && result.Type == VTJSUndefined {
						return Value{Type: VTJSUndefined}, true
					}
					return result, false
				}
				if desc.HasValue {
					return desc.Value, false
				}
			}
		}
		return Value{Type: VTJSUndefined}, false
	case VTArray:
		if target.Arr != nil && strings.EqualFold(member, "length") {
			return NewInteger(int64(len(target.Arr.Values))), false
		}
		proto := vm.jsGetIntrinsicPrototype("Array")
		if proto.Type == VTJSObject {
			desc, hasDesc := vm.jsResolveObjectMember(proto.Num, member, make(map[int64]struct{}, 4))
			if hasDesc {
				if desc.HasGetter {
					result := vm.jsCall(desc.Getter, target, nil)
					if desc.Getter.Type == VTJSFunction && len(vm.jsCallStack) > 0 && result.Type == VTJSUndefined {
						return Value{Type: VTJSUndefined}, true
					}
					return result, false
				}
				if desc.HasValue {
					return desc.Value, false
				}
			}
		}
		return Value{Type: VTJSUndefined}, false
	case VTDate:
		proto := vm.jsGetIntrinsicPrototype("Date")
		if proto.Type == VTJSObject {
			desc, hasDesc := vm.jsResolveObjectMember(proto.Num, member, make(map[int64]struct{}, 4))
			if hasDesc {
				if desc.HasGetter {
					result := vm.jsCall(desc.Getter, target, nil)
					if desc.Getter.Type == VTJSFunction && len(vm.jsCallStack) > 0 && result.Type == VTJSUndefined {
						return Value{Type: VTJSUndefined}, true
					}
					return result, false
				}
				if desc.HasValue {
					return desc.Value, false
				}
			}
		}
		return Value{Type: VTJSUndefined}, false
	case VTJSObject:
		if value, ok := vm.jsGetAliasedArgumentValue(target.Num, member); ok {
			return value, false
		}
		// ArrayBuffer property get (byteLength)
		if backing, isBuf := vm.jsArrayBuffers[target.Num]; isBuf {
			if strings.EqualFold(member, "byteLength") {
				return NewInteger(int64(len(backing))), false
			}
			if strings.EqualFold(member, "slice") {
				// Return a bound-like reference that callers can invoke; handled via jsCallMember.
				// Just return undefined — slice will be invoked via jsCallMember path.
				return Value{Type: VTJSUndefined}, false
			}
		}
		// Typed array / DataView property get
		if v, handled := vm.jsTypedArrayMemberGet(target, member); handled {
			return v, false
		}
		desc, hasDesc := vm.jsResolveObjectMember(target.Num, member, make(map[int64]struct{}, 4))
		if hasDesc {
			if desc.HasGetter {
				result := vm.jsCall(desc.Getter, target, nil)
				if desc.Getter.Type == VTJSFunction && len(vm.jsCallStack) > 0 && result.Type == VTJSUndefined {
					return Value{Type: VTJSUndefined}, true
				}
				return result, false
			}
			if desc.HasValue {
				return desc.Value, false
			}
		}
		if obj, ok := vm.jsObjectItems[target.Num]; ok {
			if val, exists := obj[member]; exists {
				return val, false
			}
		}
		return Value{Type: VTJSUndefined}, false
	case VTJSPromise:
		if strings.EqualFold(member, "then") {
			return vm.jsCreateIntrinsicFunction("Promise.prototype.then", "PromisePrototypeThen"), false
		}
		if strings.EqualFold(member, "catch") {
			return vm.jsCreateIntrinsicFunction("Promise.prototype.catch", "PromisePrototypeCatch"), false
		}
		if strings.EqualFold(member, "finally") {
			return vm.jsCreateIntrinsicFunction("Promise.prototype.finally", "PromisePrototypeFinally"), false
		}
		return Value{Type: VTJSUndefined}, false
	case VTJSFunction:
		if strings.EqualFold(member, "length") {
			return NewInteger(int64(vm.jsFunctionExpectedLength(target))), false
		}
		desc, hasDesc := vm.jsResolveObjectMember(target.Num, member, make(map[int64]struct{}, 4))
		if hasDesc {
			if desc.HasGetter {
				result := vm.jsCall(desc.Getter, target, nil)
				if desc.Getter.Type == VTJSFunction && len(vm.jsCallStack) > 0 && result.Type == VTJSUndefined {
					return Value{Type: VTJSUndefined}, true
				}
				return result, false
			}
			if desc.HasValue {
				return desc.Value, false
			}
		}
		if obj, ok := vm.jsObjectItems[target.Num]; ok {
			if val, exists := obj[member]; exists {
				return val, false
			}
		}
		return Value{Type: VTJSUndefined}, false
	default:
		return Value{Type: VTJSUndefined}, false
	}
}

func (vm *VM) jsCallMember(target Value, member string, args []Value) (Value, bool) {
	if target.Type == VTJSFunction {
		switch {
		case strings.EqualFold(member, "call"):
			callThis := Value{Type: VTJSUndefined}
			callArgs := args
			if len(args) > 0 {
				callThis = args[0]
				callArgs = args[1:]
			}
			return vm.jsCall(target, callThis, callArgs), true
		case strings.EqualFold(member, "apply"):
			applyThis := Value{Type: VTJSUndefined}
			if len(args) > 0 {
				applyThis = args[0]
			}
			applyArgs := vm.jsExtractApplyArgs(jsArgOrUndefined(args, 1))
			return vm.jsCall(target, applyThis, applyArgs), true
		case strings.EqualFold(member, "bind"):
			bindThis := jsArgOrUndefined(args, 0)
			boundArgs := []Value(nil)
			if len(args) > 1 {
				boundArgs = args[1:]
			}
			return vm.jsBindFunction(target, bindThis, boundArgs), true
		}
	}

	if target.Type == VTJSObject {
		class := vm.jsObjectStringProperty(target, "__js_type")
		switch class {
		case "Array Iterator":
			if strings.EqualFold(member, "next") {
				return vm.jsArrayIteratorNext(target), true
			}
		case "String Iterator":
			if strings.EqualFold(member, "next") {
				return vm.jsStringIteratorNext(target), true
			}
		}
	}

	if member == "slice" || member == "forEach" || member == "map" || member == "filter" || member == "at" ||
		strings.EqualFold(member, "slice") || strings.EqualFold(member, "forEach") || strings.EqualFold(member, "map") || strings.EqualFold(member, "filter") || strings.EqualFold(member, "at") {
		length, isArrayLike, deferred := vm.jsArrayLikeLength(target)
		if deferred {
			return Value{Type: VTJSUndefined}, true
		}
		if isArrayLike {
			switch {
			case strings.EqualFold(member, "at"):
				idx := 0
				if len(args) > 0 {
					idx = int(vm.jsToNumber(args[0]).Flt)
				}
				if idx < 0 {
					idx = length + idx
				}
				if idx < 0 || idx >= length {
					return Value{Type: VTJSUndefined}, true
				}
				if v, ok := vm.jsArrayLikeGetIndex(target, idx); ok {
					return v, true
				}
				return Value{Type: VTJSUndefined}, true
			case strings.EqualFold(member, "slice"):
				start := jsNormalizeRelativeIndex(int(vm.jsToNumber(jsArgOrUndefined(args, 0)).Flt), length)
				end := length
				if len(args) > 1 && args[1].Type != VTJSUndefined {
					end = jsNormalizeRelativeIndex(int(vm.jsToNumber(args[1]).Flt), length)
				}
				if end < start {
					end = start
				}
				out := make([]Value, end-start)
				for i := start; i < end; i++ {
					if v, ok := vm.jsArrayLikeGetIndex(target, i); ok {
						out[i-start] = v
					} else {
						out[i-start] = Value{Type: VTJSUndefined}
					}
				}
				return ValueFromVBArray(NewVBArrayFromValues(0, out)), true
			case strings.EqualFold(member, "forEach"):
				callback := jsArgOrUndefined(args, 0)
				if callback.Type != VTJSFunction {
					return Value{Type: VTJSUndefined}, true
				}
				thisArg := jsArgOrUndefined(args, 1)
				for i := 0; i < length; i++ {
					if !vm.jsArrayLikeHasIndex(target, i) {
						continue
					}
					item, _ := vm.jsArrayLikeGetIndex(target, i)
					_ = vm.jsCall(callback, thisArg, []Value{item, NewInteger(int64(i)), target})
				}
				return Value{Type: VTJSUndefined}, true
			case strings.EqualFold(member, "map"):
				callback := jsArgOrUndefined(args, 0)
				if callback.Type != VTJSFunction {
					return ValueFromVBArray(NewVBArrayFromValues(0, nil)), true
				}
				thisArg := jsArgOrUndefined(args, 1)
				mapped := make([]Value, length)
				for i := 0; i < length; i++ {
					if !vm.jsArrayLikeHasIndex(target, i) {
						mapped[i] = Value{Type: VTJSUndefined}
						continue
					}
					item, _ := vm.jsArrayLikeGetIndex(target, i)
					result := vm.jsCall(callback, thisArg, []Value{item, NewInteger(int64(i)), target})
					mapped[i] = result
				}
				return ValueFromVBArray(NewVBArrayFromValues(0, mapped)), true
			case strings.EqualFold(member, "filter"):
				callback := jsArgOrUndefined(args, 0)
				if callback.Type != VTJSFunction {
					return ValueFromVBArray(NewVBArrayFromValues(0, nil)), true
				}
				thisArg := jsArgOrUndefined(args, 1)
				filtered := make([]Value, 0, length)
				for i := 0; i < length; i++ {
					if !vm.jsArrayLikeHasIndex(target, i) {
						continue
					}
					item, _ := vm.jsArrayLikeGetIndex(target, i)
					result := vm.jsCall(callback, thisArg, []Value{item, NewInteger(int64(i)), target})
					if vm.jsTruthy(result) {
						filtered = append(filtered, item)
					}
				}
				return ValueFromVBArray(NewVBArrayFromValues(0, filtered)), true
			}
		}
	}

	switch target.Type {
	case VTInteger, VTDouble:
		number := vm.jsToNumber(target).Flt
		switch {
		case strings.EqualFold(member, "toString"):
			if len(args) > 0 && args[0].Type != VTJSUndefined {
				radix := int(vm.jsToNumber(args[0]).Flt)
				if radix < 2 || radix > 36 {
					return Value{Type: VTJSUndefined}, true
				}
				text := vm.jsNumberToStringRadix(number, radix)
				if !vm.jsEnsureStringSize(len(text)) || !vm.jsChargeStringWork(len(text)) {
					return Value{Type: VTJSUndefined}, true
				}
				return NewString(text), true
			}
			text := vm.jsNumberToString(number)
			if !vm.jsEnsureStringSize(len(text)) || !vm.jsChargeStringWork(len(text)) {
				return Value{Type: VTJSUndefined}, true
			}
			return NewString(text), true
		case strings.EqualFold(member, "toLocaleString"):
			text := vm.jsNumberToString(number)
			if !vm.jsEnsureStringSize(len(text)) || !vm.jsChargeStringWork(len(text)) {
				return Value{Type: VTJSUndefined}, true
			}
			return NewString(text), true
		case strings.EqualFold(member, "valueOf"):
			return NewDouble(number), true
		case strings.EqualFold(member, "toFixed"):
			digits := int(vm.jsToNumber(jsArgOrUndefined(args, 0)).Flt)
			if digits < 0 || digits > 20 {
				return Value{Type: VTJSUndefined}, true
			}
			text := vm.jsNumberToFixed(number, digits)
			if !vm.jsEnsureStringSize(len(text)) || !vm.jsChargeStringWork(len(text)) {
				return Value{Type: VTJSUndefined}, true
			}
			return NewString(text), true
		case strings.EqualFold(member, "toExponential"):
			digits := -1
			if len(args) > 0 && args[0].Type != VTJSUndefined {
				digits = int(vm.jsToNumber(args[0]).Flt)
				if digits < 0 || digits > 20 {
					return Value{Type: VTJSUndefined}, true
				}
			}
			text := vm.jsNumberToExponential(number, digits)
			if !vm.jsEnsureStringSize(len(text)) || !vm.jsChargeStringWork(len(text)) {
				return Value{Type: VTJSUndefined}, true
			}
			return NewString(text), true
		case strings.EqualFold(member, "toPrecision"):
			if len(args) == 0 || args[0].Type == VTJSUndefined {
				text := vm.jsNumberToString(number)
				if !vm.jsEnsureStringSize(len(text)) || !vm.jsChargeStringWork(len(text)) {
					return Value{Type: VTJSUndefined}, true
				}
				return NewString(text), true
			}
			precision := int(vm.jsToNumber(args[0]).Flt)
			if precision < 1 || precision > 21 {
				return Value{Type: VTJSUndefined}, true
			}
			text := vm.jsNumberToPrecision(number, precision)
			if !vm.jsEnsureStringSize(len(text)) || !vm.jsChargeStringWork(len(text)) {
				return Value{Type: VTJSUndefined}, true
			}
			return NewString(text), true
		}
	case VTString:
		text := target.Str
		runes := []rune(text)
		switch {
		case strings.EqualFold(member, "charAt"):
			idx := int(vm.jsToNumber(jsArgOrUndefined(args, 0)).Flt)
			if idx < 0 || idx >= len(runes) {
				return NewString(""), true
			}
			ch := string(runes[idx])
			if !vm.jsEnsureStringSize(len(ch)) || !vm.jsChargeStringWork(len(ch)) {
				return Value{Type: VTJSUndefined}, true
			}
			return NewString(ch), true
		case strings.EqualFold(member, "at"):
			length := len(runes)
			idx := 0
			if len(args) > 0 {
				idx = int(vm.jsToNumber(args[0]).Flt)
			}
			if idx < 0 {
				idx = length + idx
			}
			if idx < 0 || idx >= length {
				return Value{Type: VTJSUndefined}, true
			}
			ch := string(runes[idx])
			if !vm.jsEnsureStringSize(len(ch)) || !vm.jsChargeStringWork(len(ch)) {
				return Value{Type: VTJSUndefined}, true
			}
			return NewString(ch), true
		case strings.EqualFold(member, "charCodeAt"):
			idx := int(vm.jsToNumber(jsArgOrUndefined(args, 0)).Flt)
			if idx < 0 || idx >= len(runes) {
				return NewDouble(math.NaN()), true
			}
			return NewInteger(int64(runes[idx])), true
		case strings.EqualFold(member, "substring"):
			start := jsClampIndex(int(vm.jsToNumber(jsArgOrUndefined(args, 0)).Flt), len(runes))
			end := len(runes)
			if len(args) > 1 && args[1].Type != VTJSUndefined {
				end = jsClampIndex(int(vm.jsToNumber(args[1]).Flt), len(runes))
			}
			if start > end {
				start, end = end, start
			}
			out := string(runes[start:end])
			if !vm.jsEnsureStringSize(len(out)) || !vm.jsChargeStringWork(len(out)) {
				return Value{Type: VTJSUndefined}, true
			}
			return NewString(out), true
		case strings.EqualFold(member, "substr"):
			start := int(vm.jsToNumber(jsArgOrUndefined(args, 0)).Flt)
			if start < 0 {
				start = len(runes) + start
				if start < 0 {
					start = 0
				}
			}
			if start > len(runes) {
				start = len(runes)
			}
			length := len(runes) - start
			if len(args) > 1 && args[1].Type != VTJSUndefined {
				length = int(vm.jsToNumber(args[1]).Flt)
				if length < 0 {
					length = 0
				}
			}
			end := start + length
			if end > len(runes) {
				end = len(runes)
			}
			out := string(runes[start:end])
			if !vm.jsEnsureStringSize(len(out)) || !vm.jsChargeStringWork(len(out)) {
				return Value{Type: VTJSUndefined}, true
			}
			return NewString(out), true
		case strings.EqualFold(member, "slice"):
			start := jsNormalizeRelativeIndex(int(vm.jsToNumber(jsArgOrUndefined(args, 0)).Flt), len(runes))
			end := len(runes)
			if len(args) > 1 && args[1].Type != VTJSUndefined {
				end = jsNormalizeRelativeIndex(int(vm.jsToNumber(args[1]).Flt), len(runes))
			}
			if end < start {
				end = start
			}
			out := string(runes[start:end])
			if !vm.jsEnsureStringSize(len(out)) || !vm.jsChargeStringWork(len(out)) {
				return Value{Type: VTJSUndefined}, true
			}
			return NewString(out), true
		case strings.EqualFold(member, "concat"):
			total := len(text)
			parts := make([]string, 0, len(args)+1)
			parts = append(parts, text)
			for i := 0; i < len(args); i++ {
				part := vm.valueToString(args[i])
				total += len(part)
				parts = append(parts, part)
			}
			if !vm.jsEnsureStringSize(total) || !vm.jsChargeStringWork(total) {
				return Value{Type: VTJSUndefined}, true
			}
			return NewString(strings.Join(parts, "")), true
		case strings.EqualFold(member, "match"):
			if len(args) == 0 {
				return NewNull(), true
			}
			patternVal := args[0]
			pattern := vm.valueToString(patternVal)
			flags := ""
			if patternVal.Type == VTJSObject && vm.jsObjectStringProperty(patternVal, "__js_type") == "RegExp" {
				pattern = vm.jsObjectStringProperty(patternVal, "pattern")
				flags = vm.jsObjectStringProperty(patternVal, "flags")
			}
			re, err := vm.jsCompileRegExp(pattern, flags)
			if err != nil {
				return NewNull(), true
			}
			if strings.Contains(strings.ToLower(flags), "g") {
				matches := re.FindAllString(text, -1)
				if len(matches) == 0 {
					return NewNull(), true
				}
				vals := make([]Value, len(matches))
				for i := 0; i < len(matches); i++ {
					vals[i] = NewString(matches[i])
				}
				return ValueFromVBArray(NewVBArrayFromValues(0, vals)), true
			}
			first := re.FindStringSubmatch(text)
			if len(first) == 0 {
				return NewNull(), true
			}
			vals := make([]Value, len(first))
			for i := 0; i < len(first); i++ {
				vals[i] = NewString(first[i])
			}
			return ValueFromVBArray(NewVBArrayFromValues(0, vals)), true
		case strings.EqualFold(member, "search"):
			if len(args) == 0 {
				return NewInteger(-1), true
			}
			patternVal := args[0]
			pattern := vm.valueToString(patternVal)
			flags := ""
			if patternVal.Type == VTJSObject && vm.jsObjectStringProperty(patternVal, "__js_type") == "RegExp" {
				pattern = vm.jsObjectStringProperty(patternVal, "pattern")
				flags = vm.jsObjectStringProperty(patternVal, "flags")
			}
			re, err := vm.jsCompileRegExp(pattern, flags)
			if err != nil {
				return NewInteger(-1), true
			}
			loc := re.FindStringIndex(text)
			if len(loc) != 2 {
				return NewInteger(-1), true
			}
			return NewInteger(int64(loc[0])), true
		case strings.EqualFold(member, "toLowerCase"):
			out := strings.ToLower(text)
			if !vm.jsEnsureStringSize(len(out)) || !vm.jsChargeStringWork(len(out)) {
				return Value{Type: VTJSUndefined}, true
			}
			return NewString(out), true
		case strings.EqualFold(member, "toUpperCase"):
			out := strings.ToUpper(text)
			if !vm.jsEnsureStringSize(len(out)) || !vm.jsChargeStringWork(len(out)) {
				return Value{Type: VTJSUndefined}, true
			}
			return NewString(out), true
		case strings.EqualFold(member, "localeCompare"):
			other := vm.valueToString(jsArgOrUndefined(args, 0))
			cmp := strings.Compare(text, other)
			if cmp < 0 {
				return NewInteger(-1), true
			}
			if cmp > 0 {
				return NewInteger(1), true
			}
			return NewInteger(0), true
		case strings.EqualFold(member, "indexOf"):
			if len(args) == 0 {
				return NewInteger(-1), true
			}
			needle := vm.valueToString(args[0])
			start := 0
			if len(args) > 1 {
				start = int(vm.jsToNumber(args[1]).Flt)
			}
			if start < 0 {
				start = 0
			}
			if start > len(text) {
				return NewInteger(-1), true
			}
			idx := strings.Index(text[start:], needle)
			if idx < 0 {
				return NewInteger(-1), true
			}
			return NewInteger(int64(start + idx)), true
		case strings.EqualFold(member, "split"):
			if len(args) == 0 || args[0].Type == VTJSUndefined {
				return ValueFromVBArray(NewVBArrayFromValues(0, []Value{NewString(text)})), true
			}
			sep := vm.valueToString(args[0])
			var pieces []string
			if sep == "" {
				pieces = make([]string, 0, len(text))
				for _, r := range text {
					pieces = append(pieces, string(r))
				}
			} else {
				pieces = strings.Split(text, sep)
			}
			values := make([]Value, len(pieces))
			for i := range pieces {
				values[i] = NewString(pieces[i])
			}
			return ValueFromVBArray(NewVBArrayFromValues(0, values)), true
		case strings.EqualFold(member, "replace"):
			if len(args) == 0 {
				return NewString(text), true
			}
			replacementArg := jsArgOrUndefined(args, 1)
			return vm.jsStringReplace(text, args[0], replacementArg, false), true
		case strings.EqualFold(member, "replaceAll"):
			if len(args) == 0 {
				return NewString(text), true
			}
			replacementArg := jsArgOrUndefined(args, 1)
			return vm.jsStringReplace(text, args[0], replacementArg, true), true
		case strings.EqualFold(member, "trim"):
			trimmed := strings.TrimSpace(text)
			if !vm.jsEnsureStringSize(len(trimmed)) || !vm.jsChargeStringWork(len(trimmed)) {
				return Value{Type: VTJSUndefined}, true
			}
			return NewString(trimmed), true
		// ES6 String.prototype additions
		case strings.EqualFold(member, "includes"):
			if len(args) == 0 {
				return NewBool(false), true
			}
			if args[0].Type == VTJSObject && vm.jsObjectStringProperty(args[0], "__js_type") == "RegExp" {
				vm.jsThrowTypeError("First argument to String.prototype.includes must not be a RegExp")
				return Value{Type: VTJSUndefined}, true
			}
			needle := vm.valueToString(args[0])
			start := 0
			if len(args) > 1 {
				start = int(vm.jsToNumber(args[1]).Flt)
			}
			if start < 0 {
				start = 0
			}
			if start > len(text) {
				return NewBool(false), true
			}
			return NewBool(strings.Contains(text[start:], needle)), true
		case strings.EqualFold(member, "startsWith"):
			if len(args) == 0 {
				return NewBool(false), true
			}
			needle := vm.valueToString(args[0])
			start := 0
			if len(args) > 1 {
				start = int(vm.jsToNumber(args[1]).Flt)
			}
			if start < 0 {
				start = 0
			}
			if start > len(text) {
				return NewBool(false), true
			}
			return NewBool(strings.HasPrefix(text[start:], needle)), true
		case strings.EqualFold(member, "endsWith"):
			if len(args) == 0 {
				return NewBool(false), true
			}
			needle := vm.valueToString(args[0])
			end := len(text)
			if len(args) > 1 {
				e := int(vm.jsToNumber(args[1]).Flt)
				if e < end {
					end = e
				}
			}
			if end < 0 {
				end = 0
			}
			substr := text
			if end < len(text) {
				substr = text[:end]
			}
			return NewBool(strings.HasSuffix(substr, needle)), true
		case strings.EqualFold(member, "repeat"):
			if len(args) == 0 {
				return NewString(""), true
			}
			count := int(vm.jsToNumber(args[0]).Flt)
			if count < 0 {
				vm.jsThrowTypeError("Invalid count value for String.prototype.repeat")
				return Value{Type: VTJSUndefined}, true
			}
			if count == 0 || text == "" {
				return NewString(""), true
			}
			totalSize := len(text) * count
			if !vm.jsEnsureStringSize(totalSize) || !vm.jsChargeStringWork(totalSize) {
				return Value{Type: VTJSUndefined}, true
			}
			return NewString(strings.Repeat(text, count)), true
		case strings.EqualFold(member, "padStart"):
			targetLen := 0
			if len(args) > 0 {
				targetLen = int(vm.jsToNumber(args[0]).Flt)
			}
			padStr := " "
			if len(args) > 1 {
				padStr = vm.valueToString(args[1])
			}
			if targetLen <= len(text) || padStr == "" {
				return NewString(text), true
			}
			needed := targetLen - len(text)
			padded := strings.Repeat(padStr, (needed/len(padStr))+1)[:needed] + text
			if !vm.jsEnsureStringSize(len(padded)) || !vm.jsChargeStringWork(len(padded)) {
				return Value{Type: VTJSUndefined}, true
			}
			return NewString(padded), true
		case strings.EqualFold(member, "padEnd"):
			targetLen := 0
			if len(args) > 0 {
				targetLen = int(vm.jsToNumber(args[0]).Flt)
			}
			padStr := " "
			if len(args) > 1 {
				padStr = vm.valueToString(args[1])
			}
			if targetLen <= len(text) || padStr == "" {
				return NewString(text), true
			}
			needed := targetLen - len(text)
			padded := text + strings.Repeat(padStr, (needed/len(padStr))+1)[:needed]
			if !vm.jsEnsureStringSize(len(padded)) || !vm.jsChargeStringWork(len(padded)) {
				return Value{Type: VTJSUndefined}, true
			}
			return NewString(padded), true
		}
	case VTArray:
		if target.Arr == nil {
			return Value{Type: VTJSUndefined}, true
		}
		switch {
		case strings.EqualFold(member, "fill"):
			fillValue := jsArgOrUndefined(args, 0)
			length := len(target.Arr.Values)
			start := 0
			if len(args) > 1 {
				start = jsNormalizeRelativeIndex(int(vm.jsToNumber(args[1]).Flt), length)
			}
			end := length
			if len(args) > 2 && args[2].Type != VTJSUndefined {
				end = jsNormalizeRelativeIndex(int(vm.jsToNumber(args[2]).Flt), length)
			}
			if end < start {
				end = start
			}
			for i := start; i < end; i++ {
				target.Arr.Values[i] = fillValue
			}
			return target, true
		case strings.EqualFold(member, "at"):
			length := len(target.Arr.Values)
			idx := 0
			if len(args) > 0 {
				idx = int(vm.jsToNumber(args[0]).Flt)
			}
			if idx < 0 {
				idx = length + idx
			}
			if idx < 0 || idx >= length {
				return Value{Type: VTJSUndefined}, true
			}
			return target.Arr.Values[idx], true
		case strings.EqualFold(member, "flat"):
			depth := 1
			if len(args) > 0 && args[0].Type != VTJSUndefined {
				num := vm.jsToNumber(args[0])
				if math.IsInf(num.Flt, 1) {
					depth = 1000 // effectively infinity
				} else {
					depth = int(num.Flt)
				}
			}
			if depth < 0 {
				depth = 0
			}
			flattened := vm.jsArrayFlat(target.Arr.Values, depth)
			return ValueFromVBArray(NewVBArrayFromValues(0, flattened)), true
		case strings.EqualFold(member, "flatMap"):
			callback := jsArgOrUndefined(args, 0)
			if callback.Type != VTJSFunction {
				vm.jsThrowTypeError("flatMap callback must be a function")
				return Value{Type: VTJSUndefined}, true
			}
			thisArg := jsArgOrUndefined(args, 1)
			length := len(target.Arr.Values)
			out := make([]Value, 0, length)
			for i := 0; i < length; i++ {
				item := target.Arr.Values[i]
				res := vm.jsCall(callback, thisArg, []Value{item, NewInteger(int64(i)), target})
				if res.Type == VTArray && res.Arr != nil {
					out = append(out, res.Arr.Values...)
				} else {
					out = append(out, res)
				}
			}
			return ValueFromVBArray(NewVBArrayFromValues(0, out)), true
		case strings.EqualFold(member, "toSorted"):
			length := len(target.Arr.Values)
			newVals := make([]Value, length)
			copy(newVals, target.Arr.Values)
			compareFn := jsArgOrUndefined(args, 0)
			sort.Slice(newVals, func(i, j int) bool {
				a := newVals[i]
				b := newVals[j]
				if compareFn.Type == VTJSFunction {
					res := vm.jsCall(compareFn, Value{Type: VTJSUndefined}, []Value{a, b})
					return vm.jsToNumber(res).Flt < 0
				}
				return vm.valueToString(a) < vm.valueToString(b)
			})
			return ValueFromVBArray(NewVBArrayFromValues(0, newVals)), true
		case strings.EqualFold(member, "toReversed"):
			length := len(target.Arr.Values)
			newVals := make([]Value, length)
			for i := 0; i < length; i++ {
				newVals[i] = target.Arr.Values[length-1-i]
			}
			return ValueFromVBArray(NewVBArrayFromValues(0, newVals)), true
		case strings.EqualFold(member, "toSpliced"):
			length := len(target.Arr.Values)
			start := jsNormalizeRelativeIndex(int(vm.jsToNumber(jsArgOrUndefined(args, 0)).Flt), length)
			deleteCount := length - start
			if len(args) > 1 {
				deleteCount = int(vm.jsToNumber(args[1]).Flt)
				if deleteCount < 0 {
					deleteCount = 0
				}
				if deleteCount > length-start {
					deleteCount = length - start
				}
			}
			insertItems := []Value(nil)
			if len(args) > 2 {
				insertItems = args[2:]
			}
			newVals := make([]Value, 0, length-deleteCount+len(insertItems))
			newVals = append(newVals, target.Arr.Values[:start]...)
			newVals = append(newVals, insertItems...)
			newVals = append(newVals, target.Arr.Values[start+deleteCount:]...)
			return ValueFromVBArray(NewVBArrayFromValues(0, newVals)), true
		case strings.EqualFold(member, "copyWithin"):
			length := len(target.Arr.Values)
			to := jsNormalizeRelativeIndex(int(vm.jsToNumber(jsArgOrUndefined(args, 0)).Flt), length)
			from := jsNormalizeRelativeIndex(int(vm.jsToNumber(jsArgOrUndefined(args, 1)).Flt), length)
			end := length
			if len(args) > 2 && args[2].Type != VTJSUndefined {
				end = jsNormalizeRelativeIndex(int(vm.jsToNumber(args[2]).Flt), length)
			}
			if end < from {
				end = from
			}
			count := end - from
			if count > length-to {
				count = length - to
			}
			if count > 0 {
				copy(target.Arr.Values[to:to+count], target.Arr.Values[from:from+count])
			}
			return target, true
		case strings.EqualFold(member, "slice"):
			length := len(target.Arr.Values)
			start := jsNormalizeRelativeIndex(int(vm.jsToNumber(jsArgOrUndefined(args, 0)).Flt), length)
			end := length
			if len(args) > 1 && args[1].Type != VTJSUndefined {
				end = jsNormalizeRelativeIndex(int(vm.jsToNumber(args[1]).Flt), length)
			}
			if end < start {
				end = start
			}
			out := make([]Value, end-start)
			copy(out, target.Arr.Values[start:end])
			return ValueFromVBArray(NewVBArrayFromValues(0, out)), true
		case strings.EqualFold(member, "splice"):
			length := len(target.Arr.Values)
			start := jsNormalizeRelativeIndex(int(vm.jsToNumber(jsArgOrUndefined(args, 0)).Flt), length)
			deleteCount := length - start
			if len(args) > 1 {
				deleteCount = int(vm.jsToNumber(args[1]).Flt)
				if deleteCount < 0 {
					deleteCount = 0
				}
				if deleteCount > length-start {
					deleteCount = length - start
				}
			}
			removed := make([]Value, deleteCount)
			copy(removed, target.Arr.Values[start:start+deleteCount])
			insertItems := args[2:]
			newVals := make([]Value, 0, length-deleteCount+len(insertItems))
			newVals = append(newVals, target.Arr.Values[:start]...)
			newVals = append(newVals, insertItems...)
			newVals = append(newVals, target.Arr.Values[start+deleteCount:]...)
			target.Arr.Values = newVals
			return ValueFromVBArray(NewVBArrayFromValues(0, removed)), true
		case strings.EqualFold(member, "shift"):
			if len(target.Arr.Values) == 0 {
				return Value{Type: VTJSUndefined}, true
			}
			first := target.Arr.Values[0]
			target.Arr.Values = target.Arr.Values[1:]
			return first, true
		case strings.EqualFold(member, "unshift"):
			if len(args) == 0 {
				return NewInteger(int64(len(target.Arr.Values))), true
			}
			newVals := make([]Value, 0, len(args)+len(target.Arr.Values))
			newVals = append(newVals, args...)
			newVals = append(newVals, target.Arr.Values...)
			target.Arr.Values = newVals
			return NewInteger(int64(len(target.Arr.Values))), true
		case strings.EqualFold(member, "reverse"):
			for i, j := 0, len(target.Arr.Values)-1; i < j; i, j = i+1, j-1 {
				target.Arr.Values[i], target.Arr.Values[j] = target.Arr.Values[j], target.Arr.Values[i]
			}
			return target, true
		case strings.EqualFold(member, "sort"):
			compareFn := jsArgOrUndefined(args, 0)
			sort.Slice(target.Arr.Values, func(i int, j int) bool {
				a := target.Arr.Values[i]
				b := target.Arr.Values[j]
				if compareFn.Type == VTJSFunction {
					res := vm.jsCall(compareFn, Value{Type: VTJSUndefined}, []Value{a, b})
					if res.Type == VTJSUndefined {
						return vm.valueToString(a) < vm.valueToString(b)
					}
					return vm.jsToNumber(res).Flt < 0
				}
				return vm.valueToString(a) < vm.valueToString(b)
			})
			return target, true
		case strings.EqualFold(member, "concat"):
			out := make([]Value, 0, len(target.Arr.Values)+len(args))
			out = append(out, target.Arr.Values...)
			for i := 0; i < len(args); i++ {
				if args[i].Type == VTArray && args[i].Arr != nil {
					out = append(out, args[i].Arr.Values...)
					continue
				}
				if converted, ok := vm.jsAsConcatArray(args[i]); ok {
					out = append(out, converted.Arr.Values...)
					continue
				}
				out = append(out, args[i])
			}
			return ValueFromVBArray(NewVBArrayFromValues(0, out)), true
		case strings.EqualFold(member, "indexOf"):
			if len(target.Arr.Values) == 0 {
				return NewInteger(-1), true
			}
			needle := jsArgOrUndefined(args, 0)
			start := 0
			if len(args) > 1 {
				start = int(vm.jsToNumber(args[1]).Flt)
			}
			if start < 0 {
				start = len(target.Arr.Values) + start
				if start < 0 {
					start = 0
				}
			}
			for i := start; i < len(target.Arr.Values); i++ {
				if vm.jsStrictEquals(target.Arr.Values[i], needle) {
					return NewInteger(int64(i)), true
				}
			}
			return NewInteger(-1), true
		case strings.EqualFold(member, "find"):
			callback := jsArgOrUndefined(args, 0)
			if callback.Type != VTJSFunction {
				return Value{Type: VTJSUndefined}, true
			}
			thisArg := jsArgOrUndefined(args, 1)
			for i := 0; i < len(target.Arr.Values); i++ {
				result := vm.jsCall(callback, thisArg, []Value{target.Arr.Values[i], NewInteger(int64(i)), target})
				if callback.Type == VTJSFunction && len(vm.jsCallStack) > 0 && result.Type == VTJSUndefined {
					return Value{Type: VTJSUndefined}, true
				}
				if vm.jsTruthy(result) {
					return target.Arr.Values[i], true
				}
			}
			return Value{Type: VTJSUndefined}, true
		case strings.EqualFold(member, "findIndex"):
			callback := jsArgOrUndefined(args, 0)
			if callback.Type != VTJSFunction {
				return NewInteger(-1), true
			}
			thisArg := jsArgOrUndefined(args, 1)
			for i := 0; i < len(target.Arr.Values); i++ {
				result := vm.jsCall(callback, thisArg, []Value{target.Arr.Values[i], NewInteger(int64(i)), target})
				if callback.Type == VTJSFunction && len(vm.jsCallStack) > 0 && result.Type == VTJSUndefined {
					return Value{Type: VTJSUndefined}, true
				}
				if vm.jsTruthy(result) {
					return NewInteger(int64(i)), true
				}
			}
			return NewInteger(-1), true
		case strings.EqualFold(member, "lastIndexOf"):
			if len(target.Arr.Values) == 0 {
				return NewInteger(-1), true
			}
			needle := jsArgOrUndefined(args, 0)
			start := len(target.Arr.Values) - 1
			if len(args) > 1 {
				start = int(vm.jsToNumber(args[1]).Flt)
				if start >= len(target.Arr.Values) {
					start = len(target.Arr.Values) - 1
				}
			}
			if start < 0 {
				return NewInteger(-1), true
			}
			for i := start; i >= 0; i-- {
				if vm.jsStrictEquals(target.Arr.Values[i], needle) {
					return NewInteger(int64(i)), true
				}
			}
			return NewInteger(-1), true
		case strings.EqualFold(member, "push"):
			target.Arr.Values = append(target.Arr.Values, args...)
			return NewInteger(int64(len(target.Arr.Values))), true
		case strings.EqualFold(member, "__spreadPush"):
			if len(args) == 0 {
				return NewInteger(int64(len(target.Arr.Values))), true
			}
			source := args[0]
			if source.Type == VTJSUndefined || source.Type == VTNull {
				vm.jsThrowTypeError("Spread source cannot be null or undefined")
				return Value{Type: VTJSUndefined}, true
			}
			length, hasLength, deferred := vm.jsArrayLikeLength(source)
			if deferred {
				return Value{Type: VTJSUndefined}, true
			}
			if !hasLength {
				vm.jsThrowTypeError("Spread source must be array-like")
				return Value{Type: VTJSUndefined}, true
			}
			if length <= 0 {
				return NewInteger(int64(len(target.Arr.Values))), true
			}
			start := len(target.Arr.Values)
			target.Arr.Values = append(target.Arr.Values, make([]Value, length)...)
			for i := 0; i < length; i++ {
				if v, ok := vm.jsArrayLikeGetIndex(source, i); ok {
					target.Arr.Values[start+i] = v
				} else {
					target.Arr.Values[start+i] = Value{Type: VTJSUndefined}
				}
			}
			return NewInteger(int64(len(target.Arr.Values))), true
		case strings.EqualFold(member, "pop"):
			if len(target.Arr.Values) == 0 {
				return Value{Type: VTJSUndefined}, true
			}
			last := target.Arr.Values[len(target.Arr.Values)-1]
			target.Arr.Values = target.Arr.Values[:len(target.Arr.Values)-1]
			return last, true
		case strings.EqualFold(member, "join"):
			sep := ","
			if len(args) > 0 {
				sep = vm.valueToString(args[0])
			}
			if len(target.Arr.Values) == 0 {
				return NewString(""), true
			}
			parts := make([]string, len(target.Arr.Values))
			totalSize := 0
			for i := range target.Arr.Values {
				parts[i] = vm.valueToString(target.Arr.Values[i])
				totalSize += len(parts[i])
				if i > 0 {
					totalSize += len(sep)
				}
				if !vm.jsEnsureStringSize(totalSize) {
					return Value{Type: VTJSUndefined}, true
				}
			}
			if !vm.jsChargeStringWork(totalSize) {
				return Value{Type: VTJSUndefined}, true
			}
			return NewString(strings.Join(parts, sep)), true
		case strings.EqualFold(member, "forEach"):
			callback := jsArgOrUndefined(args, 0)
			if callback.Type != VTJSFunction {
				return Value{Type: VTJSUndefined}, true
			}
			thisArg := jsArgOrUndefined(args, 1)
			for i := 0; i < len(target.Arr.Values); i++ {
				result := vm.jsCall(callback, thisArg, []Value{target.Arr.Values[i], NewInteger(int64(i)), target})
				if callback.Type == VTJSFunction && len(vm.jsCallStack) > 0 && result.Type == VTJSUndefined {
					return Value{Type: VTJSUndefined}, true
				}
			}
			return Value{Type: VTJSUndefined}, true
		case strings.EqualFold(member, "every"):
			callback := jsArgOrUndefined(args, 0)
			if callback.Type != VTJSFunction {
				return NewBool(false), true
			}
			thisArg := jsArgOrUndefined(args, 1)
			for i := 0; i < len(target.Arr.Values); i++ {
				result := vm.jsCall(callback, thisArg, []Value{target.Arr.Values[i], NewInteger(int64(i)), target})
				if callback.Type == VTJSFunction && len(vm.jsCallStack) > 0 && result.Type == VTJSUndefined {
					return Value{Type: VTJSUndefined}, true
				}
				if !vm.jsTruthy(result) {
					return NewBool(false), true
				}
			}
			return NewBool(true), true
		case strings.EqualFold(member, "some"):
			callback := jsArgOrUndefined(args, 0)
			if callback.Type != VTJSFunction {
				return NewBool(false), true
			}
			thisArg := jsArgOrUndefined(args, 1)
			for i := 0; i < len(target.Arr.Values); i++ {
				result := vm.jsCall(callback, thisArg, []Value{target.Arr.Values[i], NewInteger(int64(i)), target})
				if callback.Type == VTJSFunction && len(vm.jsCallStack) > 0 && result.Type == VTJSUndefined {
					return Value{Type: VTJSUndefined}, true
				}
				if vm.jsTruthy(result) {
					return NewBool(true), true
				}
			}
			return NewBool(false), true
		case strings.EqualFold(member, "map"):
			callback := jsArgOrUndefined(args, 0)
			if callback.Type != VTJSFunction {
				return ValueFromVBArray(NewVBArrayFromValues(0, nil)), true
			}
			thisArg := jsArgOrUndefined(args, 1)
			mapped := make([]Value, len(target.Arr.Values))
			for i := 0; i < len(target.Arr.Values); i++ {
				result := vm.jsCall(callback, thisArg, []Value{target.Arr.Values[i], NewInteger(int64(i)), target})
				if callback.Type == VTJSFunction && len(vm.jsCallStack) > 0 && result.Type == VTJSUndefined {
					return Value{Type: VTJSUndefined}, true
				}
				mapped[i] = result
			}
			return ValueFromVBArray(NewVBArrayFromValues(0, mapped)), true
		case strings.EqualFold(member, "filter"):
			callback := jsArgOrUndefined(args, 0)
			if callback.Type != VTJSFunction {
				return ValueFromVBArray(NewVBArrayFromValues(0, nil)), true
			}
			thisArg := jsArgOrUndefined(args, 1)
			filtered := make([]Value, 0, len(target.Arr.Values))
			for i := 0; i < len(target.Arr.Values); i++ {
				result := vm.jsCall(callback, thisArg, []Value{target.Arr.Values[i], NewInteger(int64(i)), target})
				if callback.Type == VTJSFunction && len(vm.jsCallStack) > 0 && result.Type == VTJSUndefined {
					return Value{Type: VTJSUndefined}, true
				}
				if vm.jsTruthy(result) {
					filtered = append(filtered, target.Arr.Values[i])
				}
			}
			return ValueFromVBArray(NewVBArrayFromValues(0, filtered)), true
		case strings.EqualFold(member, "reduce"):
			if len(target.Arr.Values) == 0 && len(args) < 2 {
				return Value{Type: VTJSUndefined}, true
			}
			callback := jsArgOrUndefined(args, 0)
			if callback.Type != VTJSFunction {
				return Value{Type: VTJSUndefined}, true
			}
			acc := Value{Type: VTJSUndefined}
			start := 0
			if len(args) > 1 {
				acc = args[1]
			} else {
				acc = target.Arr.Values[0]
				start = 1
			}
			for i := start; i < len(target.Arr.Values); i++ {
				result := vm.jsCall(callback, Value{Type: VTJSUndefined}, []Value{acc, target.Arr.Values[i], NewInteger(int64(i)), target})
				if callback.Type == VTJSFunction && len(vm.jsCallStack) > 0 && result.Type == VTJSUndefined {
					return Value{Type: VTJSUndefined}, true
				}
				acc = result
			}
			return acc, true
		case strings.EqualFold(member, "reduceRight"):
			if len(target.Arr.Values) == 0 && len(args) < 2 {
				return Value{Type: VTJSUndefined}, true
			}
			callback := jsArgOrUndefined(args, 0)
			if callback.Type != VTJSFunction {
				return Value{Type: VTJSUndefined}, true
			}
			acc := Value{Type: VTJSUndefined}
			start := len(target.Arr.Values) - 1
			if len(args) > 1 {
				acc = args[1]
			} else {
				acc = target.Arr.Values[start]
				start--
			}
			for i := start; i >= 0; i-- {
				result := vm.jsCall(callback, Value{Type: VTJSUndefined}, []Value{acc, target.Arr.Values[i], NewInteger(int64(i)), target})
				if callback.Type == VTJSFunction && len(vm.jsCallStack) > 0 && result.Type == VTJSUndefined {
					return Value{Type: VTJSUndefined}, true
				}
				acc = result
			}
			return acc, true
		}
	case VTDate:
		switch {
		case strings.EqualFold(member, "getFullYear"):
			t := valueToTimeInLocale(vm, target).In(builtinCurrentLocation(vm))
			return NewInteger(int64(t.Year())), true
		case strings.EqualFold(member, "getMonth"):
			t := valueToTimeInLocale(vm, target).In(builtinCurrentLocation(vm))
			return NewInteger(int64(int(t.Month()) - 1)), true
		case strings.EqualFold(member, "getDate"):
			t := valueToTimeInLocale(vm, target).In(builtinCurrentLocation(vm))
			return NewInteger(int64(t.Day())), true
		case strings.EqualFold(member, "getDay"):
			t := valueToTimeInLocale(vm, target).In(builtinCurrentLocation(vm))
			return NewInteger(int64(t.Weekday())), true
		case strings.EqualFold(member, "getHours"):
			t := valueToTimeInLocale(vm, target).In(builtinCurrentLocation(vm))
			return NewInteger(int64(t.Hour())), true
		case strings.EqualFold(member, "getMinutes"):
			t := valueToTimeInLocale(vm, target).In(builtinCurrentLocation(vm))
			return NewInteger(int64(t.Minute())), true
		case strings.EqualFold(member, "getSeconds"):
			t := valueToTimeInLocale(vm, target).In(builtinCurrentLocation(vm))
			return NewInteger(int64(t.Second())), true
		case strings.EqualFold(member, "getMilliseconds"):
			t := valueToTimeInLocale(vm, target).In(builtinCurrentLocation(vm))
			return NewInteger(int64(t.Nanosecond() / int(time.Millisecond))), true
		case strings.EqualFold(member, "getTimezoneOffset"):
			t := valueToTimeInLocale(vm, target).In(builtinCurrentLocation(vm))
			_, offsetSeconds := t.Zone()
			return NewInteger(int64(-(offsetSeconds / 60))), true
		case strings.EqualFold(member, "getTime"):
			t := valueToTimeInLocale(vm, target)
			return NewInteger(t.UnixNano() / int64(time.Millisecond)), true
		case strings.EqualFold(member, "now"):
			return NewInteger(time.Now().UnixNano() / int64(time.Millisecond)), true
		case strings.EqualFold(member, "toString"):
			t := valueToTimeInLocale(vm, target).In(builtinCurrentLocation(vm))
			return NewString(t.Format("Mon Jan 02 2006 15:04:05 GMT-0700")), true
		case strings.EqualFold(member, "toLocaleString"):
			return NewString(vm.dateToLocalizedString(target)), true
		case strings.EqualFold(member, "toUTCString"):
			t := valueToTimeInLocale(vm, target).UTC()
			return NewString(t.Format(time.RFC1123)), true
		case strings.EqualFold(member, "toISOString"):
			t := valueToTimeInLocale(vm, target).UTC()
			return NewString(t.Format(time.RFC3339)), true
		case strings.EqualFold(member, "toJSON"):
			result, _ := vm.jsCallMember(target, "toISOString", nil)
			return result, true
		case strings.EqualFold(member, "valueOf"):
			t := valueToTimeInLocale(vm, target)
			return NewInteger(t.UnixNano() / int64(time.Millisecond)), true
		}
	case VTJSPromise:
		switch {
		case strings.EqualFold(member, "then"):
			return vm.jsPromiseThen(target, args), true
		case strings.EqualFold(member, "catch"):
			return vm.jsPromiseCatch(target, args), true
		case strings.EqualFold(member, "finally"):
			return vm.jsPromiseFinally(target, args), true
		}
	case VTJSObject:
		switch {
		case strings.EqualFold(member, "hasOwnProperty"):
			return NewBool(vm.jsObjectHasOwnProperty(target, vm.valueToString(jsArgOrUndefined(args, 0)))), true
		case strings.EqualFold(member, "propertyIsEnumerable"):
			return NewBool(vm.jsObjectPropertyIsEnumerable(target, vm.valueToString(jsArgOrUndefined(args, 0)))), true
		case strings.EqualFold(member, "isPrototypeOf"):
			return NewBool(vm.jsObjectIsPrototypeOf(target, jsArgOrUndefined(args, 0))), true
		case strings.EqualFold(member, "toString"):
			return NewString(vm.jsObjectToStringTag(target)), true
		case strings.EqualFold(member, "toLocaleString"):
			return NewString(vm.jsObjectToStringTag(target)), true
		case strings.EqualFold(member, "valueOf"):
			return target, true
		}
		objType := vm.jsObjectStringProperty(target, "__js_type")
		if objType == "" {
			objType = vm.jsObjectStringProperty(target, "__js_ctor")
		}
		switch objType {
		case "Date":
			switch {
			case strings.EqualFold(member, "now"):
				return NewInteger(time.Now().UnixNano() / int64(time.Millisecond)), true
			case strings.EqualFold(member, "parse"):
				if len(args) == 0 {
					return NewDouble(math.NaN()), true
				}
				t := valueToTimeInLocale(vm, args[0])
				if t.IsZero() {
					return NewDouble(math.NaN()), true
				}
				return NewInteger(t.UnixNano() / int64(time.Millisecond)), true
			case strings.EqualFold(member, "UTC"):
				year := int(vm.jsToNumber(jsArgOrUndefined(args, 0)).Flt)
				month := int(vm.jsToNumber(jsArgOrUndefined(args, 1)).Flt)
				day := 1
				hour := 0
				minute := 0
				second := 0
				millisecond := 0
				if len(args) > 2 {
					day = int(vm.jsToNumber(args[2]).Flt)
				}
				if len(args) > 3 {
					hour = int(vm.jsToNumber(args[3]).Flt)
				}
				if len(args) > 4 {
					minute = int(vm.jsToNumber(args[4]).Flt)
				}
				if len(args) > 5 {
					second = int(vm.jsToNumber(args[5]).Flt)
				}
				if len(args) > 6 {
					millisecond = int(vm.jsToNumber(args[6]).Flt)
				}
				t := time.Date(year, time.Month(month+1), day, hour, minute, second, millisecond*int(time.Millisecond), time.UTC)
				return NewInteger(t.UnixNano() / int64(time.Millisecond)), true
			}
		case "Array":
			switch {
			case strings.EqualFold(member, "isArray"):
				if len(args) == 0 {
					return NewBool(false), true
				}
				return NewBool(args[0].Type == VTArray), true
			case strings.EqualFold(member, "of"):
				if len(args) == 0 {
					return ValueFromVBArray(NewVBArrayFromValues(0, nil)), true
				}
				vals := make([]Value, len(args))
				copy(vals, args)
				return ValueFromVBArray(NewVBArrayFromValues(0, vals)), true
			case strings.EqualFold(member, "from"):
				if len(args) == 0 || args[0].Type == VTJSUndefined || args[0].Type == VTNull {
					vm.jsThrowTypeError("Array.from requires an array-like source")
					return Value{Type: VTJSUndefined}, true
				}
				source := args[0]
				switch source.Type {
				case VTArray:
					if source.Arr == nil {
						return ValueFromVBArray(NewVBArrayFromValues(0, nil)), true
					}
					vals := make([]Value, len(source.Arr.Values))
					copy(vals, source.Arr.Values)
					return ValueFromVBArray(NewVBArrayFromValues(0, vals)), true
				case VTString:
					r := []rune(source.Str)
					vals := make([]Value, len(r))
					for i := 0; i < len(r); i++ {
						vals[i] = NewString(string(r[i]))
					}
					return ValueFromVBArray(NewVBArrayFromValues(0, vals)), true
				default:
					length, hasLength, deferred := vm.jsArrayLikeLength(source)
					if deferred {
						return Value{Type: VTJSUndefined}, true
					}
					if !hasLength {
						vm.jsThrowTypeError("Array.from source is not array-like")
						return Value{Type: VTJSUndefined}, true
					}
					vals := make([]Value, length)
					for i := 0; i < length; i++ {
						if v, ok := vm.jsArrayLikeGetIndex(source, i); ok {
							vals[i] = v
						} else {
							vals[i] = Value{Type: VTJSUndefined}
						}
					}
					return ValueFromVBArray(NewVBArrayFromValues(0, vals)), true
				}
			}
		case "Object":
			switch {
			case strings.EqualFold(member, "create"):
				objID := vm.allocJSID()
				obj := make(map[string]Value, 8)
				vm.jsObjectItems[objID] = obj
				vm.jsPropertyItems[objID] = make(map[string]jsPropertyDescriptor, 8)
				if len(args) > 0 {
					if args[0].Type == VTJSObject || args[0].Type == VTNull {
						obj["__js_proto"] = args[0]
					}
				}
				created := Value{Type: VTJSObject, Num: objID}
				if len(args) > 1 {
					_, _ = vm.jsCallMember(target, "defineProperties", []Value{created, args[1]})
				}
				return created, true
			case strings.EqualFold(member, "getPrototypeOf"):
				if len(args) == 0 || args[0].Type != VTJSObject {
					return NewNull(), true
				}
				if obj, ok := vm.jsObjectItems[args[0].Num]; ok {
					if proto, exists := obj["__js_proto"]; exists {
						return proto, true
					}
				}
				return NewNull(), true
			case strings.EqualFold(member, "fromEntries"):
				if len(args) == 0 || args[0].Type == VTJSUndefined || args[0].Type == VTNull {
					vm.jsThrowTypeError("Object.fromEntries requires an iterable")
					return Value{Type: VTJSUndefined}, true
				}
				iterable := args[0]
				objID := vm.allocJSID()
				obj := make(map[string]Value, 8)
				vm.jsObjectItems[objID] = obj
				vm.jsPropertyItems[objID] = make(map[string]jsPropertyDescriptor, 8)
				if iterable.Type == VTArray && iterable.Arr != nil {
					for _, entry := range iterable.Arr.Values {
						if entry.Type == VTArray && entry.Arr != nil && len(entry.Arr.Values) >= 2 {
							key := vm.valueToString(entry.Arr.Values[0])
							val := entry.Arr.Values[1]
							obj[key] = val
							vm.jsSetDescriptor(objID, key, jsDefaultPropertyDescriptor(val))
						}
					}
				}
				return Value{Type: VTJSObject, Num: objID}, true
			case strings.EqualFold(member, "keys"):
				if len(args) == 0 || args[0].Type == VTJSUndefined || args[0].Type == VTNull {
					vm.jsThrowTypeError("Object.keys called on null or undefined")
					return Value{Type: VTJSUndefined}, true
				}
				keys := make([]string, 0)
				source := args[0]
				switch source.Type {
				case VTJSObject, VTJSFunction:
					keys = vm.jsObjectOwnEnumerableKeys(source.Num)
				case VTArray:
					if source.Arr != nil {
						keys = make([]string, len(source.Arr.Values))
						for i := 0; i < len(source.Arr.Values); i++ {
							keys[i] = strconv.Itoa(i)
						}
					}
				case VTString:
					runes := []rune(source.Str)
					keys = make([]string, len(runes))
					for i := 0; i < len(runes); i++ {
						keys[i] = strconv.Itoa(i)
					}
				}
				values := make([]Value, len(keys))
				for i := 0; i < len(keys); i++ {
					values[i] = NewString(keys[i])
				}
				return ValueFromVBArray(NewVBArrayFromValues(0, values)), true
			case strings.EqualFold(member, "values"):
				if len(args) == 0 || args[0].Type == VTJSUndefined || args[0].Type == VTNull {
					vm.jsThrowTypeError("Object.values called on null or undefined")
					return Value{Type: VTJSUndefined}, true
				}
				source := args[0]
				values := make([]Value, 0)
				switch source.Type {
				case VTJSObject, VTJSFunction:
					keys := vm.jsObjectOwnEnumerableKeys(source.Num)
					values = make([]Value, len(keys))
					for i := 0; i < len(keys); i++ {
						v, deferred := vm.jsMemberGet(source, keys[i])
						if deferred {
							return Value{Type: VTJSUndefined}, true
						}
						values[i] = v
					}
				case VTArray:
					if source.Arr != nil {
						values = make([]Value, len(source.Arr.Values))
						copy(values, source.Arr.Values)
					}
				case VTString:
					runes := []rune(source.Str)
					values = make([]Value, len(runes))
					for i := 0; i < len(runes); i++ {
						values[i] = NewString(string(runes[i]))
					}
				}
				return ValueFromVBArray(NewVBArrayFromValues(0, values)), true
			case strings.EqualFold(member, "entries"):
				if len(args) == 0 || args[0].Type == VTJSUndefined || args[0].Type == VTNull {
					vm.jsThrowTypeError("Object.entries called on null or undefined")
					return Value{Type: VTJSUndefined}, true
				}
				source := args[0]
				keys := make([]string, 0)
				vals := make([]Value, 0)
				switch source.Type {
				case VTJSObject, VTJSFunction:
					keys = vm.jsObjectOwnEnumerableKeys(source.Num)
					vals = make([]Value, len(keys))
					for i := 0; i < len(keys); i++ {
						v, deferred := vm.jsMemberGet(source, keys[i])
						if deferred {
							return Value{Type: VTJSUndefined}, true
						}
						vals[i] = v
					}
				case VTArray:
					if source.Arr != nil {
						keys = make([]string, len(source.Arr.Values))
						vals = make([]Value, len(source.Arr.Values))
						for i := 0; i < len(source.Arr.Values); i++ {
							keys[i] = strconv.Itoa(i)
							vals[i] = source.Arr.Values[i]
						}
					}
				case VTString:
					runes := []rune(source.Str)
					keys = make([]string, len(runes))
					vals = make([]Value, len(runes))
					for i := 0; i < len(runes); i++ {
						keys[i] = strconv.Itoa(i)
						vals[i] = NewString(string(runes[i]))
					}
				}
				entries := make([]Value, len(keys))
				for i := 0; i < len(keys); i++ {
					entries[i] = ValueFromVBArray(NewVBArrayFromValues(0, []Value{NewString(keys[i]), vals[i]}))
				}
				return ValueFromVBArray(NewVBArrayFromValues(0, entries)), true
			case strings.EqualFold(member, "assign"):
				if len(args) == 0 || args[0].Type == VTJSUndefined || args[0].Type == VTNull {
					vm.jsThrowTypeError("Object.assign target cannot be null or undefined")
					return Value{Type: VTJSUndefined}, true
				}
				targetObj := args[0]
				if targetObj.Type != VTJSObject && targetObj.Type != VTJSFunction && targetObj.Type != VTArray {
					vm.jsThrowTypeError("Object.assign target must be an object")
					return Value{Type: VTJSUndefined}, true
				}
				for si := 1; si < len(args); si++ {
					source := args[si]
					if source.Type == VTJSUndefined || source.Type == VTNull {
						continue
					}
					switch source.Type {
					case VTJSObject, VTJSFunction:
						keys := vm.jsObjectOwnEnumerableKeys(source.Num)
						for i := 0; i < len(keys); i++ {
							v, deferred := vm.jsMemberGet(source, keys[i])
							if deferred {
								return Value{Type: VTJSUndefined}, true
							}
							if targetObj.Type == VTArray {
								if idx, err := strconv.Atoi(keys[i]); err == nil {
									vm.jsIndexSet(targetObj, NewInteger(int64(idx)), v)
								}
							} else {
								vm.jsMemberSet(targetObj, keys[i], v)
							}
						}
					case VTArray:
						if source.Arr == nil {
							continue
						}
						for i := 0; i < len(source.Arr.Values); i++ {
							key := strconv.Itoa(i)
							if targetObj.Type == VTArray {
								vm.jsIndexSet(targetObj, NewInteger(int64(i)), source.Arr.Values[i])
							} else {
								vm.jsMemberSet(targetObj, key, source.Arr.Values[i])
							}
						}
					case VTString:
						runes := []rune(source.Str)
						for i := 0; i < len(runes); i++ {
							v := NewString(string(runes[i]))
							if targetObj.Type == VTArray {
								vm.jsIndexSet(targetObj, NewInteger(int64(i)), v)
							} else {
								vm.jsMemberSet(targetObj, strconv.Itoa(i), v)
							}
						}
					}
				}
				return targetObj, true
			case strings.EqualFold(member, "getOwnPropertyNames"):
				if len(args) == 0 {
					return ValueFromVBArray(NewVBArrayFromValues(0, nil)), true
				}
				names := vm.jsObjectOwnPropertyNames(args[0])
				values := make([]Value, len(names))
				for i := 0; i < len(names); i++ {
					values[i] = NewString(names[i])
				}
				return ValueFromVBArray(NewVBArrayFromValues(0, values)), true
			case strings.EqualFold(member, "defineProperty"):
				if len(args) < 3 || (args[0].Type != VTJSObject && args[0].Type != VTJSFunction) {
					return jsArgOrUndefined(args, 0), true
				}
				objID := args[0].Num
				name := vm.valueToString(args[1])
				current, currentExists := vm.jsGetDescriptor(objID, name)
				if !currentExists && !vm.jsObjectIsExtensible(args[0]) {
					return args[0], true
				}
				spec := vm.jsReadDefinePropertySpec(args[2])
				if !vm.jsValidateDefinePropertyTransition(current, currentExists, spec) {
					return args[0], true
				}
				finalDesc := vm.jsApplyDefinePropertySpec(current, currentExists, spec)
				vm.jsSetDescriptor(objID, name, finalDesc)
				return args[0], true
			case strings.EqualFold(member, "defineProperties"):
				if len(args) < 2 || (args[0].Type != VTJSObject && args[0].Type != VTJSFunction) || args[1].Type != VTJSObject {
					return jsArgOrUndefined(args, 0), true
				}
				descObj := vm.jsObjectItems[args[1].Num]
				for key, oneDesc := range descObj {
					if strings.HasPrefix(key, jsInternalPropPrefix) || oneDesc.Type != VTJSObject {
						continue
					}
					_, _ = vm.jsCallMember(target, "defineProperty", []Value{args[0], NewString(key), oneDesc})
				}
				return args[0], true
			case strings.EqualFold(member, "getOwnPropertyDescriptor"):
				if len(args) < 2 || args[0].Type == VTJSUndefined || args[0].Type == VTNull {
					vm.jsThrowTypeError("Object.getOwnPropertyDescriptor called on null or undefined")
					return Value{Type: VTJSUndefined}, true
				}
				if args[0].Type != VTJSObject && args[0].Type != VTJSFunction {
					return Value{Type: VTJSUndefined}, true
				}
				name := vm.jsPropertyKeyFromValue(args[1])
				desc, ok := vm.jsGetDescriptor(args[0].Num, name)
				if !ok {
					return Value{Type: VTJSUndefined}, true
				}
				return vm.jsBuildDescriptorObject(desc), true
			case strings.EqualFold(member, "getOwnPropertyDescriptors"):
				if len(args) < 1 || args[0].Type == VTJSUndefined || args[0].Type == VTNull {
					vm.jsThrowTypeError("Object.getOwnPropertyDescriptors called on null or undefined")
					return Value{Type: VTJSUndefined}, true
				}
				if args[0].Type != VTJSObject && args[0].Type != VTJSFunction {
					return Value{Type: VTJSUndefined}, true
				}
				resultID := vm.allocJSID()
				resultObj := make(map[string]Value, 8)
				vm.jsObjectItems[resultID] = resultObj
				vm.jsPropertyItems[resultID] = make(map[string]jsPropertyDescriptor, 8)
				names := vm.jsObjectOwnPropertyNames(args[0])
				for i := 0; i < len(names); i++ {
					desc, ok := vm.jsGetDescriptor(args[0].Num, names[i])
					if !ok {
						continue
					}
					descObj := vm.jsBuildDescriptorObject(desc)
					resultObj[names[i]] = descObj
					vm.jsSetDescriptor(resultID, names[i], jsDefaultPropertyDescriptor(descObj))
				}
				return Value{Type: VTJSObject, Num: resultID}, true
			case strings.EqualFold(member, "preventExtensions"):
				if len(args) == 0 {
					return Value{Type: VTJSUndefined}, true
				}
				if args[0].Type == VTJSObject || args[0].Type == VTJSFunction {
					vm.jsSetObjectExtensible(args[0].Num, false)
				}
				return args[0], true
			case strings.EqualFold(member, "isExtensible"):
				if len(args) == 0 {
					return NewBool(false), true
				}
				return NewBool(vm.jsObjectIsExtensible(args[0])), true
			case strings.EqualFold(member, "seal"):
				if len(args) == 0 {
					return Value{Type: VTJSUndefined}, true
				}
				return vm.jsObjectSeal(args[0]), true
			case strings.EqualFold(member, "isSealed"):
				if len(args) == 0 {
					return NewBool(false), true
				}
				return NewBool(vm.jsObjectIsSealed(args[0])), true
			case strings.EqualFold(member, "freeze"):
				if len(args) == 0 {
					return Value{Type: VTJSUndefined}, true
				}
				return vm.jsObjectFreeze(args[0]), true
			case strings.EqualFold(member, "isFrozen"):
				if len(args) == 0 {
					return NewBool(false), true
				}
				return NewBool(vm.jsObjectIsFrozen(args[0])), true
			}
		case "JSON":
			switch {
			case strings.EqualFold(member, "parse"):
				if len(args) == 0 {
					return Value{Type: VTJSUndefined}, true
				}
				return vm.jsJSONParse(vm.valueToString(args[0])), true
			case strings.EqualFold(member, "stringify"):
				if len(args) == 0 {
					return NewString("null"), true
				}
				jsonText := vm.jsJSONStringify(args[0])
				if !vm.jsEnsureStringSize(len(jsonText)) || !vm.jsChargeStringWork(len(jsonText)) {
					return Value{Type: VTJSUndefined}, true
				}
				return NewString(jsonText), true
			}
		case "Math":
			switch {
			case strings.EqualFold(member, "abs"):
				if len(args) == 0 {
					return NewDouble(math.NaN()), true
				}
				return NewDouble(math.Abs(vm.jsToNumber(args[0]).Flt)), true
			case strings.EqualFold(member, "sin"):
				return NewDouble(math.Sin(vm.jsToNumber(jsArgOrUndefined(args, 0)).Flt)), true
			case strings.EqualFold(member, "cos"):
				return NewDouble(math.Cos(vm.jsToNumber(jsArgOrUndefined(args, 0)).Flt)), true
			case strings.EqualFold(member, "tan"):
				return NewDouble(math.Tan(vm.jsToNumber(jsArgOrUndefined(args, 0)).Flt)), true
			case strings.EqualFold(member, "asin"):
				return NewDouble(math.Asin(vm.jsToNumber(jsArgOrUndefined(args, 0)).Flt)), true
			case strings.EqualFold(member, "acos"):
				return NewDouble(math.Acos(vm.jsToNumber(jsArgOrUndefined(args, 0)).Flt)), true
			case strings.EqualFold(member, "atan"):
				return NewDouble(math.Atan(vm.jsToNumber(jsArgOrUndefined(args, 0)).Flt)), true
			case strings.EqualFold(member, "atan2"):
				y := vm.jsToNumber(jsArgOrUndefined(args, 0)).Flt
				x := vm.jsToNumber(jsArgOrUndefined(args, 1)).Flt
				return NewDouble(math.Atan2(y, x)), true
			case strings.EqualFold(member, "ceil"):
				return NewDouble(math.Ceil(vm.jsToNumber(jsArgOrUndefined(args, 0)).Flt)), true
			case strings.EqualFold(member, "floor"):
				return NewDouble(math.Floor(vm.jsToNumber(jsArgOrUndefined(args, 0)).Flt)), true
			case strings.EqualFold(member, "trunc"):
				return NewDouble(math.Trunc(vm.jsToNumber(jsArgOrUndefined(args, 0)).Flt)), true
			case strings.EqualFold(member, "sign"):
				n := vm.jsToNumber(jsArgOrUndefined(args, 0)).Flt
				if math.IsNaN(n) {
					return NewDouble(math.NaN()), true
				}
				if n > 0 {
					return NewInteger(1), true
				}
				if n < 0 {
					return NewInteger(-1), true
				}
				return NewInteger(0), true
			case strings.EqualFold(member, "cbrt"):
				return NewDouble(math.Cbrt(vm.jsToNumber(jsArgOrUndefined(args, 0)).Flt)), true
			case strings.EqualFold(member, "round"):
				return NewDouble(math.Round(vm.jsToNumber(jsArgOrUndefined(args, 0)).Flt)), true
			case strings.EqualFold(member, "sqrt"):
				return NewDouble(math.Sqrt(vm.jsToNumber(jsArgOrUndefined(args, 0)).Flt)), true
			case strings.EqualFold(member, "pow"):
				base := vm.jsToNumber(jsArgOrUndefined(args, 0)).Flt
				exp := vm.jsToNumber(jsArgOrUndefined(args, 1)).Flt
				return NewDouble(math.Pow(base, exp)), true
			case strings.EqualFold(member, "max"):
				if len(args) == 0 {
					return NewDouble(math.Inf(-1)), true
				}
				maxVal := vm.jsToNumber(args[0]).Flt
				for i := 1; i < len(args); i++ {
					n := vm.jsToNumber(args[i]).Flt
					if n > maxVal || math.IsNaN(n) {
						maxVal = n
					}
				}
				return NewDouble(maxVal), true
			case strings.EqualFold(member, "min"):
				if len(args) == 0 {
					return NewDouble(math.Inf(1)), true
				}
				minVal := vm.jsToNumber(args[0]).Flt
				for i := 1; i < len(args); i++ {
					n := vm.jsToNumber(args[i]).Flt
					if n < minVal || math.IsNaN(n) {
						minVal = n
					}
				}
				return NewDouble(minVal), true
			case strings.EqualFold(member, "random"):
				return NewDouble(rand.Float64()), true
			}
		case "Set":
			switch {
			case strings.EqualFold(member, "add"):
				if store, ok := vm.jsSetItems[target.Num]; ok {
					arg := jsArgOrUndefined(args, 0)
					store[vm.jsValueMapKey(arg)] = arg
				}
				return target, true
			case strings.EqualFold(member, "has"):
				store, ok := vm.jsSetItems[target.Num]
				if !ok {
					return NewBool(false), true
				}
				_, exists := store[vm.jsValueMapKey(jsArgOrUndefined(args, 0))]
				return NewBool(exists), true
			case strings.EqualFold(member, "delete"):
				store, ok := vm.jsSetItems[target.Num]
				if !ok {
					return NewBool(false), true
				}
				key := vm.jsValueMapKey(jsArgOrUndefined(args, 0))
				_, exists := store[key]
				if exists {
					delete(store, key)
				}
				return NewBool(exists), true
			case strings.EqualFold(member, "clear"):
				if store, ok := vm.jsSetItems[target.Num]; ok {
					clear(store)
				}
				return Value{Type: VTJSUndefined}, true
			}
		case "Map":
			switch {
			case strings.EqualFold(member, "set"):
				if store, ok := vm.jsMapItems[target.Num]; ok {
					store[vm.jsValueMapKey(jsArgOrUndefined(args, 0))] = jsArgOrUndefined(args, 1)
				}
				return target, true
			case strings.EqualFold(member, "has"):
				store, ok := vm.jsMapItems[target.Num]
				if !ok {
					return NewBool(false), true
				}
				_, exists := store[vm.jsValueMapKey(jsArgOrUndefined(args, 0))]
				return NewBool(exists), true
			case strings.EqualFold(member, "delete"):
				store, ok := vm.jsMapItems[target.Num]
				if !ok {
					return NewBool(false), true
				}
				key := vm.jsValueMapKey(jsArgOrUndefined(args, 0))
				_, exists := store[key]
				if exists {
					delete(store, key)
				}
				return NewBool(exists), true
			case strings.EqualFold(member, "clear"):
				if store, ok := vm.jsMapItems[target.Num]; ok {
					clear(store)
				}
				return Value{Type: VTJSUndefined}, true
			}
		case "Number":
			// ES6 Number static methods.
			switch {
			case strings.EqualFold(member, "isInteger"):
				if len(args) == 0 {
					return NewBool(false), true
				}
				v := args[0]
				switch v.Type {
				case VTInteger:
					return NewBool(true), true
				case VTDouble:
					return NewBool(!math.IsNaN(v.Flt) && !math.IsInf(v.Flt, 0) && math.Trunc(v.Flt) == v.Flt), true
				default:
					return NewBool(false), true
				}
			case strings.EqualFold(member, "isNaN"):
				if len(args) == 0 {
					return NewBool(false), true
				}
				v := args[0]
				// ES6 Number.isNaN does NOT coerce; only actual NaN returns true.
				if v.Type == VTDouble {
					return NewBool(math.IsNaN(v.Flt)), true
				}
				return NewBool(false), true
			case strings.EqualFold(member, "isFinite"):
				if len(args) == 0 {
					return NewBool(false), true
				}
				v := args[0]
				// ES6 Number.isFinite does NOT coerce; only actual numbers are checked.
				switch v.Type {
				case VTInteger:
					return NewBool(true), true
				case VTDouble:
					return NewBool(!math.IsNaN(v.Flt) && !math.IsInf(v.Flt, 0)), true
				default:
					return NewBool(false), true
				}
			case strings.EqualFold(member, "isSafeInteger"):
				if len(args) == 0 {
					return NewBool(false), true
				}
				v := args[0]
				switch v.Type {
				case VTInteger:
					return NewBool(v.Num >= -9007199254740991 && v.Num <= 9007199254740991), true
				case VTDouble:
					return NewBool(!math.IsNaN(v.Flt) && !math.IsInf(v.Flt, 0) && math.Trunc(v.Flt) == v.Flt && v.Flt >= -9007199254740991 && v.Flt <= 9007199254740991), true
				default:
					return NewBool(false), true
				}
			case strings.EqualFold(member, "parseInt"):
				return vm.jsParseIntES5(args), true
			case strings.EqualFold(member, "parseFloat"):
				return vm.jsParseFloatES5(args), true
			}
		case "RegExp":
			if strings.EqualFold(member, "test") {
				pattern := vm.jsObjectStringProperty(target, "pattern")
				flags := vm.jsObjectStringProperty(target, "flags")
				needle := ""
				if len(args) > 0 {
					needle = vm.valueToString(args[0])
				}
				re, err := vm.jsCompileRegExp(pattern, flags)
				if err != nil {
					return NewBool(false), true
				}
				return NewBool(re.MatchString(needle)), true
			}
		case "Enumerator":
			switch {
			case strings.EqualFold(member, "atEnd"):
				return NewBool(vm.jsEnumeratorAtEnd(target)), true
			case strings.EqualFold(member, "moveNext"):
				vm.jsEnumeratorMoveNext(target)
				return Value{Type: VTJSUndefined}, true
			case strings.EqualFold(member, "moveFirst"):
				vm.jsEnumeratorMoveFirst(target)
				return Value{Type: VTJSUndefined}, true
			case strings.EqualFold(member, "item"):
				return vm.jsEnumeratorItem(target), true
			}
		case "VBArray":
			switch {
			case strings.EqualFold(member, "dimensions"):
				return NewInteger(int64(vm.jsVBArrayDimensions(target))), true
			case strings.EqualFold(member, "lbound"):
				dim := 1
				if len(args) > 0 {
					dim = int(vm.jsToNumber(args[0]).Flt)
				}
				lower, _, ok := arrayBounds(vm.jsVBArraySource(target), dim)
				if !ok {
					return Value{Type: VTJSUndefined}, true
				}
				return NewInteger(int64(lower)), true
			case strings.EqualFold(member, "ubound"):
				dim := 1
				if len(args) > 0 {
					dim = int(vm.jsToNumber(args[0]).Flt)
				}
				_, upper, ok := arrayBounds(vm.jsVBArraySource(target), dim)
				if !ok {
					return Value{Type: VTJSUndefined}, true
				}
				return NewInteger(int64(upper)), true
			case strings.EqualFold(member, "toArray"):
				return vm.jsVBArrayToJSArray(vm.jsVBArraySource(target)), true
			case strings.EqualFold(member, "getItem"):
				return vm.jsVBArrayGetItem(target, args), true
			}
		case "Symbol":
			// Symbol constructor static methods
			switch {
			case strings.EqualFold(member, "for"):
				key := ""
				if len(args) > 0 {
					key = vm.valueToString(args[0])
				}
				return vm.jsSymbolFor(key), true
			case strings.EqualFold(member, "keyFor"):
				return vm.jsSymbolKeyFor(jsArgOrUndefined(args, 0)), true
			}
		case "ArrayBuffer":
			// ArrayBuffer.isView(arg) static method (called on the ctor object)
			if strings.EqualFold(member, "isView") {
				if len(args) == 0 {
					return NewBool(false), true
				}
				if args[0].Type != VTJSObject {
					return NewBool(false), true
				}
				t := vm.jsObjectStringProperty(args[0], "__js_type")
				return NewBool(jsIsTypedArrayType(t)), true
			}
			// ArrayBuffer instance method: slice
			if strings.EqualFold(member, "slice") {
				return vm.jsArrayBufferSlice(target, args), true
			}
		default:
			// Instance methods on typed arrays and DataView
			if jsIsTypedArrayType(objType) {
				switch {
				case strings.EqualFold(member, "set"):
					return vm.jsTypedArraySet(target, args), true
				case strings.EqualFold(member, "subarray"):
					return vm.jsTypedArraySubarray(target, args), true
				case strings.EqualFold(member, "fill"):
					return vm.jsTypedArrayFill(target, args), true
				case strings.EqualFold(member, "slice"):
					// Returns a new typed array copy of the range
					return vm.jsTypedArraySubarray(target, args), true
				case strings.EqualFold(member, "forEach"):
					length := 0
					if items, ok := vm.jsObjectItems[target.Num]; ok {
						if lv, ok2 := items["__js_byte_length"]; ok2 {
							elemSz := jsTypedArrayElementSize(objType)
							if elemSz > 0 {
								length = int(lv.Num) / elemSz
							}
						}
					}
					callback := jsArgOrUndefined(args, 0)
					if callback.Type == VTJSFunction {
						buf, byteOffset, _, elemSize, ok := vm.jsGetTypedArrayInfo(target)
						if ok {
							for i := 0; i < length; i++ {
								v := jsReadTypedArrayElement(objType, elemSize, buf, byteOffset, i)
								_ = vm.jsCall(callback, Value{Type: VTJSUndefined}, []Value{v, NewInteger(int64(i)), target})
							}
						}
					}
					return Value{Type: VTJSUndefined}, true
				case strings.EqualFold(member, "indexOf"):
					target2 := jsArgOrUndefined(args, 0)
					fromIndex := 0
					if len(args) > 1 {
						fromIndex = int(vm.jsToNumber(args[1]).Flt)
					}
					buf, byteOffset, byteLength, elemSize, ok := vm.jsGetTypedArrayInfo(target)
					if !ok {
						return NewInteger(-1), true
					}
					length := byteLength / elemSize
					searchVal := vm.jsToNumber(target2).Flt
					for i := fromIndex; i < length; i++ {
						v := jsReadTypedArrayElement(objType, elemSize, buf, byteOffset, i)
						if vm.jsToNumber(v).Flt == searchVal {
							return NewInteger(int64(i)), true
						}
					}
					return NewInteger(-1), true
				default:
					// DataView method calls
					if objType == "DataView" {
						if v, handled := vm.jsDataViewCallMember(target, member, args); handled {
							return v, true
						}
					}
				}
			}
		}
	}

	// General fallback: dispatch based on target type
	switch target.Type {
	case VTNativeObject:
		return vm.dispatchNativeCall(target.Num, member, args), true
	case VTJSObject, VTJSFunction, VTArray, VTString, VTDate:
		callee, deferred := vm.jsMemberGet(target, member)
		if deferred {
			return Value{Type: VTJSUndefined}, true
		}
		if callee.Type != VTJSUndefined {
			return vm.jsCall(callee, target, args), true
		}
	}

	return Value{Type: VTJSUndefined}, false
}

// jsNumberToString formats a numeric primitive using JScript-compatible defaults.
func (vm *VM) jsNumberToString(n float64) string {
	if math.IsNaN(n) {
		return "NaN"
	}
	if math.IsInf(n, 1) {
		return "Infinity"
	}
	if math.IsInf(n, -1) {
		return "-Infinity"
	}
	if n == 0 {
		return "0"
	}
	return vm.jsNormalizeExponent(strconv.FormatFloat(n, 'g', -1, 64))
}

// jsNumberToStringRadix formats integral numbers in the requested radix.
func (vm *VM) jsNumberToStringRadix(n float64, radix int) string {
	if math.IsNaN(n) || math.IsInf(n, 0) {
		return vm.jsNumberToString(n)
	}
	if n == 0 {
		return "0"
	}
	truncated := math.Trunc(n)
	if truncated != n {
		return vm.jsNumberToString(n)
	}
	sign := ""
	if truncated < 0 {
		sign = "-"
		truncated = -truncated
	}
	const digits = "0123456789abcdefghijklmnopqrstuvwxyz"
	value := uint64(truncated)
	buf := [65]byte{}
	idx := len(buf)
	for value > 0 {
		idx--
		buf[idx] = digits[value%uint64(radix)]
		value /= uint64(radix)
	}
	if idx == len(buf) {
		return "0"
	}
	return sign + string(buf[idx:])
}

// jsNumberToFixed implements Number.prototype.toFixed with bounded precision.
func (vm *VM) jsNumberToFixed(n float64, digits int) string {
	if math.IsNaN(n) || math.IsInf(n, 0) {
		return vm.jsNumberToString(n)
	}
	if n >= 1e21 || n <= -1e21 {
		return vm.jsNumberToString(n)
	}
	if n == 0 {
		n = 0
	}
	return string(ftoa.FToStr(n, ftoa.ModeFixed, digits, nil))
}

// jsNumberToExponential implements Number.prototype.toExponential.
func (vm *VM) jsNumberToExponential(n float64, digits int) string {
	if math.IsNaN(n) || math.IsInf(n, 0) {
		return vm.jsNumberToString(n)
	}
	if n == 0 {
		n = 0
	}
	if digits < 0 {
		return vm.jsNormalizeExponent(strconv.FormatFloat(n, 'e', -1, 64))
	}
	return vm.jsNormalizeExponent(strconv.FormatFloat(n, 'e', digits, 64))
}

// jsNumberToPrecision implements Number.prototype.toPrecision.
func (vm *VM) jsNumberToPrecision(n float64, precision int) string {
	if math.IsNaN(n) || math.IsInf(n, 0) {
		return vm.jsNumberToString(n)
	}
	if n == 0 {
		n = 0
	}
	return vm.jsNormalizeExponent(strconv.FormatFloat(n, 'g', precision, 64))
}

// jsNormalizeExponent removes zero-padding in exponent suffix to match JS formatting.
func (vm *VM) jsNormalizeExponent(text string) string {
	ePos := strings.IndexAny(text, "eE")
	if ePos < 0 || ePos+2 >= len(text) {
		return text
	}
	base := text[:ePos]
	exp := text[ePos+1:]
	sign := exp[0:1]
	digits := strings.TrimLeft(exp[1:], "0")
	if digits == "" {
		digits = "0"
	}
	return base + "e" + sign + digits
}

// jsArgOrUndefined returns args[idx] or undefined for missing arguments.
func jsArgOrUndefined(args []Value, idx int) Value {
	if idx >= 0 && idx < len(args) {
		return args[idx]
	}
	return Value{Type: VTJSUndefined}
}

// jsStringReplace implements String.prototype.replace and replaceAll with size guards.
func (vm *VM) jsStringReplace(source string, patternArg Value, replacementArg Value, replaceAll bool) Value {
	if patternArg.Type == VTJSObject {
		objType := vm.jsObjectStringProperty(patternArg, "__js_type")
		if objType == "RegExp" {
			pattern := vm.jsObjectStringProperty(patternArg, "pattern")
			flags := vm.jsObjectStringProperty(patternArg, "flags")
			return vm.jsStringReplaceRegex(source, pattern, flags, replacementArg, replaceAll)
		}
	}

	search := vm.valueToString(patternArg)
	useCallback := replacementArg.Type == VTJSFunction
	replacement := vm.valueToString(replacementArg)
	if search == "" {
		if replaceAll {
			var b strings.Builder
			for i := 0; i <= len(source); i++ {
				repl := replacement
				if useCallback {
					cb := vm.jsCall(replacementArg, Value{Type: VTJSUndefined}, []Value{NewString(""), NewInteger(int64(i)), NewString(source)})
					repl = vm.valueToString(cb)
				}
				b.WriteString(repl)
				if i < len(source) {
					b.WriteByte(source[i])
				}
			}
			out := b.String()
			if !vm.jsEnsureStringSize(len(out)) || !vm.jsChargeStringWork(len(out)) {
				return Value{Type: VTJSUndefined}
			}
			return NewString(out)
		}
		repl := replacement
		if useCallback {
			cb := vm.jsCall(replacementArg, Value{Type: VTJSUndefined}, []Value{NewString(""), NewInteger(0), NewString(source)})
			repl = vm.valueToString(cb)
		}
		out := repl + source
		if !vm.jsEnsureStringSize(len(out)) || !vm.jsChargeStringWork(len(out)) {
			return Value{Type: VTJSUndefined}
		}
		return NewString(out)
	}

	if !replaceAll {
		idx := strings.Index(source, search)
		if idx < 0 {
			if !vm.jsEnsureStringSize(len(source)) || !vm.jsChargeStringWork(len(source)) {
				return Value{Type: VTJSUndefined}
			}
			return NewString(source)
		}
		repl := replacement
		if useCallback {
			cb := vm.jsCall(replacementArg, Value{Type: VTJSUndefined}, []Value{NewString(search), NewInteger(int64(idx)), NewString(source)})
			repl = vm.valueToString(cb)
		}
		out := source[:idx] + repl + source[idx+len(search):]
		if !vm.jsEnsureStringSize(len(out)) || !vm.jsChargeStringWork(len(out)) {
			return Value{Type: VTJSUndefined}
		}
		return NewString(out)
	}

	var b strings.Builder
	start := 0
	for {
		idx := strings.Index(source[start:], search)
		if idx < 0 {
			b.WriteString(source[start:])
			break
		}
		idx += start
		b.WriteString(source[start:idx])
		repl := replacement
		if useCallback {
			cb := vm.jsCall(replacementArg, Value{Type: VTJSUndefined}, []Value{NewString(search), NewInteger(int64(idx)), NewString(source)})
			repl = vm.valueToString(cb)
		}
		b.WriteString(repl)
		start = idx + len(search)
	}
	out := b.String()
	if !vm.jsEnsureStringSize(len(out)) || !vm.jsChargeStringWork(len(out)) {
		return Value{Type: VTJSUndefined}
	}
	return NewString(out)
}

// jsStringReplaceRegex applies one or all RegExp matches to the source string.
func (vm *VM) jsStringReplaceRegex(source string, pattern string, flags string, replacementArg Value, replaceAll bool) Value {
	rePattern := pattern
	flagsLower := strings.ToLower(flags)
	if strings.Contains(flagsLower, "i") {
		rePattern = "(?i)" + rePattern
	}
	re, err := regexp.Compile(rePattern)
	if err != nil {
		return NewString(source)
	}
	replacement := vm.valueToString(replacementArg)
	useCallback := replacementArg.Type == VTJSFunction
	useAll := replaceAll || strings.Contains(flagsLower, "g")
	limit := 1
	if useAll {
		limit = -1
	}
	matches := re.FindAllStringSubmatchIndex(source, limit)
	if len(matches) == 0 {
		if !vm.jsEnsureStringSize(len(source)) || !vm.jsChargeStringWork(len(source)) {
			return Value{Type: VTJSUndefined}
		}
		return NewString(source)
	}
	var b strings.Builder
	last := 0
	for i := 0; i < len(matches); i++ {
		idx := matches[i]
		if len(idx) < 2 {
			continue
		}
		start := idx[0]
		end := idx[1]
		if start < last {
			continue
		}
		b.WriteString(source[last:start])
		repl := replacement
		if useCallback {
			captureCount := (len(idx) / 2) - 1
			callbackArgs := make([]Value, 0, captureCount+3)
			callbackArgs = append(callbackArgs, NewString(source[start:end]))
			for c := 0; c < captureCount; c++ {
				capStart := idx[2+(c*2)]
				capEnd := idx[3+(c*2)]
				if capStart >= 0 && capEnd >= capStart {
					callbackArgs = append(callbackArgs, NewString(source[capStart:capEnd]))
				} else {
					callbackArgs = append(callbackArgs, Value{Type: VTJSUndefined})
				}
			}
			callbackArgs = append(callbackArgs, NewInteger(int64(start)), NewString(source))
			cb := vm.jsCall(replacementArg, Value{Type: VTJSUndefined}, callbackArgs)
			repl = vm.valueToString(cb)
		}
		b.WriteString(repl)
		last = end
	}
	b.WriteString(source[last:])
	out := b.String()
	if !vm.jsEnsureStringSize(len(out)) || !vm.jsChargeStringWork(len(out)) {
		return Value{Type: VTJSUndefined}
	}
	return NewString(out)
}

// jsEnumeratorSource returns the wrapped collection for a JScript Enumerator.
func (vm *VM) jsEnumeratorSource(obj Value) Value {
	if obj.Type != VTJSObject {
		return Value{Type: VTJSUndefined}
	}
	items, ok := vm.jsObjectItems[obj.Num]
	if !ok {
		return Value{Type: VTJSUndefined}
	}
	source, ok := items["__js_enum_source"]
	if !ok {
		return Value{Type: VTJSUndefined}
	}
	return source
}

// jsEnumeratorIndex returns the current Enumerator cursor index.
func (vm *VM) jsEnumeratorIndex(obj Value) int {
	if obj.Type != VTJSObject {
		return 0
	}
	items, ok := vm.jsObjectItems[obj.Num]
	if !ok {
		return 0
	}
	v, ok := items["__js_enum_index"]
	if !ok {
		return 0
	}
	return int(vm.jsToNumber(v).Flt)
}

// jsEnumeratorSetIndex writes the Enumerator cursor index.
func (vm *VM) jsEnumeratorSetIndex(obj Value, idx int) {
	if obj.Type != VTJSObject {
		return
	}
	items, ok := vm.jsObjectItems[obj.Num]
	if !ok {
		return
	}
	items["__js_enum_index"] = NewInteger(int64(idx))
}

// jsEnumeratorItemCount returns the number of available items in an Enumerator.
func (vm *VM) jsEnumeratorItemCount(obj Value) int {
	source := vm.jsEnumeratorSource(obj)
	switch source.Type {
	case VTArray:
		if source.Arr == nil {
			return 0
		}
		return len(source.Arr.Values)
	case VTJSObject:
		items, ok := vm.jsObjectItems[obj.Num]
		if !ok {
			return 0
		}
		keyList, ok := items["__js_enum_keys"]
		if !ok || keyList.Type != VTArray || keyList.Arr == nil {
			return 0
		}
		return len(keyList.Arr.Values)
	case VTNativeObject:
		if dict, ok := vm.dictionaryItems[source.Num]; ok && dict != nil {
			return len(dict.keys)
		}
		countVal := vm.dispatchMemberGet(source, "Count")
		return int(vm.jsToNumber(countVal).Flt)
	default:
		return 0
	}
}

// jsNativeCollectionItemByKeyIndex resolves one collection value from Key(index) lookups.
func (vm *VM) jsNativeCollectionItemByKeyIndex(source Value, idx int) Value {
	if idx < 0 {
		return Value{Type: VTJSUndefined}
	}
	keyMethod := vm.dispatchMemberGet(source, "Key")
	if keyMethod.Type != VTNativeObject {
		return Value{Type: VTJSUndefined}
	}
	lookup := func(pos int) Value {
		key := vm.dispatchNativeCall(keyMethod.Num, "", []Value{NewInteger(int64(pos))})
		if key.Type == VTJSUndefined || key.Type == VTEmpty {
			return Value{Type: VTJSUndefined}
		}
		keyStr := vm.valueToString(key)
		if keyStr == "" {
			return Value{Type: VTJSUndefined}
		}
		return vm.dispatchNativeCall(source.Num, "", []Value{NewString(keyStr)})
	}
	value := lookup(idx + 1)
	if value.Type != VTJSUndefined && value.Type != VTEmpty {
		return value
	}
	return lookup(idx)
}

// jsEnumeratorAtEnd reports whether the Enumerator cursor reached the end.
func (vm *VM) jsEnumeratorAtEnd(obj Value) bool {
	return vm.jsEnumeratorIndex(obj) >= vm.jsEnumeratorItemCount(obj)
}

// jsEnumeratorMoveNext advances the Enumerator cursor by one.
func (vm *VM) jsEnumeratorMoveNext(obj Value) {
	vm.jsEnumeratorSetIndex(obj, vm.jsEnumeratorIndex(obj)+1)
}

// jsEnumeratorMoveFirst resets the Enumerator cursor to the first item.
func (vm *VM) jsEnumeratorMoveFirst(obj Value) {
	vm.jsEnumeratorSetIndex(obj, 0)
}

// jsEnumeratorItem returns the current item in the wrapped collection.
func (vm *VM) jsEnumeratorItem(obj Value) Value {
	source := vm.jsEnumeratorSource(obj)
	idx := vm.jsEnumeratorIndex(obj)
	switch source.Type {
	case VTArray:
		if source.Arr == nil || idx < 0 || idx >= len(source.Arr.Values) {
			return Value{Type: VTJSUndefined}
		}
		return source.Arr.Values[idx]
	case VTJSObject:
		items, ok := vm.jsObjectItems[obj.Num]
		if !ok {
			return Value{Type: VTJSUndefined}
		}
		keyList, ok := items["__js_enum_keys"]
		if !ok || keyList.Type != VTArray || keyList.Arr == nil || idx < 0 || idx >= len(keyList.Arr.Values) {
			return Value{Type: VTJSUndefined}
		}
		key := keyList.Arr.Values[idx].Str
		objMap, ok := vm.jsObjectItems[source.Num]
		if !ok {
			return Value{Type: VTJSUndefined}
		}
		if v, exists := objMap[key]; exists {
			return v
		}
		return Value{Type: VTJSUndefined}
	case VTNativeObject:
		if idx < 0 {
			return Value{Type: VTJSUndefined}
		}
		if dict, ok := vm.dictionaryItems[source.Num]; ok && dict != nil {
			if idx >= len(dict.keys) {
				return Value{Type: VTJSUndefined}
			}
			return dict.keys[idx]
		}
		zeroBased := vm.dispatchNativeCall(source.Num, "Item", []Value{NewInteger(int64(idx))})
		if zeroBased.Type != VTJSUndefined && zeroBased.Type != VTEmpty {
			return zeroBased
		}
		oneBased := vm.dispatchNativeCall(source.Num, "Item", []Value{NewInteger(int64(idx + 1))})
		if oneBased.Type != VTJSUndefined && oneBased.Type != VTEmpty {
			return oneBased
		}
		return vm.jsNativeCollectionItemByKeyIndex(source, idx)
	default:
		return Value{Type: VTJSUndefined}
	}
}

// jsVBArrayDepth returns the maximum nested array depth starting at one VBArray node.
func (vm *VM) jsVBArrayDepth(arr *VBArray) int {
	if arr == nil {
		return 0
	}
	maxChildDepth := 0
	for i := 0; i < len(arr.Values); i++ {
		child, ok := toVBArray(arr.Values[i])
		if !ok {
			continue
		}
		depth := vm.jsVBArrayDepth(child)
		if depth > maxChildDepth {
			maxChildDepth = depth
		}
	}
	return 1 + maxChildDepth
}

// jsVBArraySource extracts the wrapped VBArray value from a JScript VBArray wrapper object.
func (vm *VM) jsVBArraySource(obj Value) Value {
	if obj.Type != VTJSObject {
		return Value{Type: VTJSUndefined}
	}
	items, ok := vm.jsObjectItems[obj.Num]
	if !ok {
		return Value{Type: VTJSUndefined}
	}
	source, ok := items["__js_vbarray_source"]
	if !ok {
		return Value{Type: VTJSUndefined}
	}
	return source
}

// jsVBArrayDimensions returns the number of dimensions for the wrapped VBArray.
func (vm *VM) jsVBArrayDimensions(obj Value) int {
	value := vm.jsVBArraySource(obj)
	arr, ok := toVBArray(value)
	if !ok {
		return 0
	}
	return vm.jsVBArrayDepth(arr)
}

// jsVBArrayToJSArray converts a VBArray value into a zero-based array value.
func (vm *VM) jsVBArrayToJSArray(value Value) Value {
	arr, ok := toVBArray(value)
	if !ok || arr == nil {
		return Value{Type: VTJSUndefined}
	}
	converted := make([]Value, len(arr.Values))
	for i := 0; i < len(arr.Values); i++ {
		if child, childOK := toVBArray(arr.Values[i]); childOK {
			converted[i] = vm.jsVBArrayToJSArray(ValueFromVBArray(child))
		} else {
			converted[i] = arr.Values[i]
		}
	}
	return ValueFromVBArray(NewVBArrayFromValues(0, converted))
}

// jsVBArrayGetItem fetches one element from VBArray using one or more indexes.
func (vm *VM) jsVBArrayGetItem(obj Value, args []Value) Value {
	current := vm.jsVBArraySource(obj)
	if current.Type != VTArray || current.Arr == nil {
		return Value{Type: VTJSUndefined}
	}
	if len(args) == 0 {
		return Value{Type: VTJSUndefined}
	}
	for i := 0; i < len(args); i++ {
		if current.Type != VTArray || current.Arr == nil {
			return Value{Type: VTJSUndefined}
		}
		idx := int(vm.jsToNumber(args[i]).Flt)
		offset := idx - current.Arr.Lower
		if offset < 0 || offset >= len(current.Arr.Values) {
			return Value{Type: VTJSUndefined}
		}
		current = current.Arr.Values[offset]
	}
	return current
}

func (vm *VM) jsMemberSet(target Value, member string, val Value) {
	if target.Type == VTJSUninitialized {
		vm.jsThrowReferenceError("Must call super constructor in derived class before accessing 'this'")
		return
	}
	switch target.Type {
	case VTNativeObject:
		vm.dispatchMemberSet(target.Num, member, val)
	case VTJSObject, VTJSFunction:
		var targetID int64
		if target.Type == VTJSObject || target.Type == VTJSFunction {
			targetID = target.Num
		}
		if vm.jsSetAliasedArgumentValue(targetID, member, val) {
			return
		}
		// Typed array: handle numeric-index writes via string key.
		if vm.jsTypedArrayMemberSet(target, member, val) {
			return
		}
		if strings.HasPrefix(member, jsAccessorGetterPrefix) {
			name := strings.TrimPrefix(member, jsAccessorGetterPrefix)
			desc, _ := vm.jsGetDescriptor(targetID, name)
			desc.Getter = val
			desc.HasGetter = true
			desc.HasValue = false
			desc.Writable = false
			desc.Enumerable = true
			desc.Configurable = true
			vm.jsSetDescriptor(targetID, name, desc)
			return
		}
		if strings.HasPrefix(member, jsAccessorSetterPrefix) {
			name := strings.TrimPrefix(member, jsAccessorSetterPrefix)
			desc, _ := vm.jsGetDescriptor(targetID, name)
			desc.Setter = val
			desc.HasSetter = true
			desc.HasValue = false
			desc.Writable = false
			desc.Enumerable = true
			desc.Configurable = true
			vm.jsSetDescriptor(targetID, name, desc)
			return
		}

		desc, hasDesc := vm.jsGetDescriptor(targetID, member)
		if hasDesc {
			if desc.HasSetter {
				vm.jsCall(desc.Setter, target, []Value{val})
				return
			}
			if !desc.Writable {
				// In strict mode, writing to a non-writable property is a TypeError.
				if vm.jsStrictMode {
					vm.jsThrowTypeError(fmt.Sprintf("Cannot assign to read-only property '%s'", member))
				}
				return
			}
			desc.Value = val
			desc.HasValue = true
			vm.jsSetDescriptor(targetID, member, desc)
			return
		}

		if proto := vm.jsGetPrototypeValue(target); proto.Type == VTJSObject {
			if inherited, ok := vm.jsResolveObjectMember(proto.Num, member, make(map[int64]struct{}, 4)); ok {
				if inherited.HasSetter {
					vm.jsCall(inherited.Setter, target, []Value{val})
					return
				}
				if !inherited.Writable {
					if vm.jsStrictMode {
						vm.jsThrowTypeError(fmt.Sprintf("Cannot assign to read-only property '%s'", member))
					}
					return
				}
			}
		}

		obj, ok := vm.jsObjectItems[targetID]
		if !ok {
			obj = make(map[string]Value, 8)
			vm.jsObjectItems[targetID] = obj
		}
		if !vm.jsObjectIsExtensible(target) {
			// In strict mode, adding a property to a non-extensible object is a TypeError.
			if vm.jsStrictMode {
				vm.jsThrowTypeError(fmt.Sprintf("Cannot add property '%s', object is not extensible", member))
			}
			return
		}
		obj[member] = val
		vm.jsSetDescriptor(targetID, member, jsDefaultPropertyDescriptor(val))
	}
}

func jsDescriptorIsAccessor(desc jsPropertyDescriptor) bool {
	return desc.HasGetter || desc.HasSetter
}

func jsDescriptorIsData(desc jsPropertyDescriptor) bool {
	return desc.HasValue
}

func (vm *VM) jsReadDefinePropertySpec(descVal Value) jsDefinePropertySpec {
	spec := jsDefinePropertySpec{desc: jsPropertyDescriptor{Enumerable: false, Configurable: false, Writable: false}}
	if descVal.Type != VTJSObject {
		return spec
	}
	dObj, ok := vm.jsObjectItems[descVal.Num]
	if !ok {
		return spec
	}
	if v, has := dObj["value"]; has {
		spec.desc.Value = v
		spec.desc.HasValue = true
	}
	if v, has := dObj["get"]; has {
		spec.desc.Getter = v
		spec.desc.HasGetter = true
	}
	if v, has := dObj["set"]; has {
		spec.desc.Setter = v
		spec.desc.HasSetter = true
	}
	if v, has := dObj["enumerable"]; has {
		spec.hasEnumerable = true
		spec.desc.Enumerable = vm.jsToDescriptorBoolean(v, false)
	}
	if v, has := dObj["configurable"]; has {
		spec.hasConfigurable = true
		spec.desc.Configurable = vm.jsToDescriptorBoolean(v, false)
	}
	if v, has := dObj["writable"]; has {
		spec.hasWritable = true
		spec.desc.Writable = vm.jsToDescriptorBoolean(v, false)
	}
	return spec
}

func (vm *VM) jsBuildDescriptorObject(desc jsPropertyDescriptor) Value {
	descObjID := vm.allocJSID()
	descObj := make(map[string]Value, 8)
	if desc.HasValue {
		descObj["value"] = desc.Value
		descObj["writable"] = NewBool(desc.Writable)
	}
	if desc.HasGetter {
		descObj["get"] = desc.Getter
	} else if desc.HasSetter {
		descObj["get"] = Value{Type: VTJSUndefined}
	}
	if desc.HasSetter {
		descObj["set"] = desc.Setter
	} else if desc.HasGetter {
		descObj["set"] = Value{Type: VTJSUndefined}
	}
	descObj["enumerable"] = NewBool(desc.Enumerable)
	descObj["configurable"] = NewBool(desc.Configurable)
	vm.jsObjectItems[descObjID] = descObj
	vm.jsPropertyItems[descObjID] = make(map[string]jsPropertyDescriptor, 8)
	for k, v := range descObj {
		vm.jsSetDescriptor(descObjID, k, jsDefaultPropertyDescriptor(v))
	}
	return Value{Type: VTJSObject, Num: descObjID}
}

func (vm *VM) jsValidateDefinePropertyTransition(current jsPropertyDescriptor, currentExists bool, spec jsDefinePropertySpec) bool {
	if spec.desc.HasValue && (spec.desc.HasGetter || spec.desc.HasSetter) {
		return false
	}
	if spec.hasWritable && (spec.desc.HasGetter || spec.desc.HasSetter) {
		return false
	}
	if !currentExists {
		return true
	}
	if current.Configurable {
		return true
	}
	if spec.hasConfigurable && spec.desc.Configurable {
		return false
	}
	if spec.hasEnumerable && spec.desc.Enumerable != current.Enumerable {
		return false
	}
	isCurrentAccessor := jsDescriptorIsAccessor(current)
	isSpecAccessor := spec.desc.HasGetter || spec.desc.HasSetter
	isSpecData := spec.desc.HasValue || spec.hasWritable
	if isSpecAccessor && jsDescriptorIsData(current) {
		return false
	}
	if isSpecData && isCurrentAccessor {
		return false
	}
	if isCurrentAccessor {
		if spec.desc.HasGetter && !vm.jsStrictEquals(spec.desc.Getter, current.Getter) {
			return false
		}
		if spec.desc.HasSetter && !vm.jsStrictEquals(spec.desc.Setter, current.Setter) {
			return false
		}
		return true
	}
	if !current.Writable {
		if spec.hasWritable && spec.desc.Writable {
			return false
		}
		if spec.desc.HasValue && !vm.jsStrictEquals(spec.desc.Value, current.Value) {
			return false
		}
	}
	return true
}

func (vm *VM) jsApplyDefinePropertySpec(current jsPropertyDescriptor, currentExists bool, spec jsDefinePropertySpec) jsPropertyDescriptor {
	result := current
	if !currentExists {
		result = jsPropertyDescriptor{Enumerable: false, Configurable: false, Writable: false}
	}
	isSpecAccessor := spec.desc.HasGetter || spec.desc.HasSetter
	isSpecData := spec.desc.HasValue || spec.hasWritable
	if isSpecAccessor {
		result.HasValue = false
		result.Writable = false
		result.Value = Value{Type: VTJSUndefined}
		if spec.desc.HasGetter {
			result.Getter = spec.desc.Getter
			result.HasGetter = true
		}
		if spec.desc.HasSetter {
			result.Setter = spec.desc.Setter
			result.HasSetter = true
		}
	} else if isSpecData {
		if jsDescriptorIsAccessor(result) {
			result.Getter = Value{Type: VTJSUndefined}
			result.Setter = Value{Type: VTJSUndefined}
			result.HasGetter = false
			result.HasSetter = false
		}
		if !currentExists {
			result.HasValue = true
			result.Value = Value{Type: VTJSUndefined}
		}
		if spec.desc.HasValue {
			result.HasValue = true
			result.Value = spec.desc.Value
		}
		if spec.hasWritable {
			result.Writable = spec.desc.Writable
		}
	}
	if spec.hasEnumerable {
		result.Enumerable = spec.desc.Enumerable
	}
	if spec.hasConfigurable {
		result.Configurable = spec.desc.Configurable
	}
	if !currentExists && !isSpecAccessor {
		if !result.HasValue {
			result.HasValue = true
			result.Value = Value{Type: VTJSUndefined}
		}
	}
	return result
}

func (vm *VM) jsEnumerateForInKeys(source Value) []string {
	if source.Type == VTJSObject {
		return vm.jsObjectOwnEnumerableKeys(source.Num)
	}

	if source.Type == VTArray && source.Arr != nil && len(source.Arr.Values) > 0 {
		keys := make([]string, 0, len(source.Arr.Values))
		for i := 0; i < len(source.Arr.Values); i++ {
			keys = append(keys, strconv.Itoa(source.Arr.Lower+i))
		}
		return keys
	}

	return nil
}

// jsEnumerateForOfValues collects the iterable values from a source for for...of traversal.
// Arrays yield their elements, Strings yield each character, JS Sets yield their members,
// JS Maps yield [key, value] pair arrays, and JS Arrays (VTJSObject with __js_class Array)
// yield indexed elements. All other types produce an empty slice.
func (vm *VM) jsEnumerateForOfValues(source Value) []Value {
	// Fast path: VBScript/internal Array — yield elements directly.
	if source.Type == VTArray && source.Arr != nil {
		out := make([]Value, len(source.Arr.Values))
		copy(out, source.Arr.Values)
		return out
	}

	// String — yield each character as a single-character string.
	if source.Type == VTString {
		runes := []rune(source.Str)
		out := make([]Value, len(runes))
		for i, r := range runes {
			out[i] = NewString(string(r))
		}
		return out
	}

	// JS Object — inspect the runtime class to decide how to iterate.
	if source.Type == VTJSObject {
		class := vm.jsObjectStringProperty(source, "__js_class")
		if class == "" {
			class = vm.jsObjectStringProperty(source, "__js_type")
		}
		if class == "" {
			class = vm.jsObjectStringProperty(source, "__js_ctor")
		}
		switch class {
		case "Array":
			// JS Array: iterate by numeric index up to .length.
			lengthVal, _ := vm.jsMemberGet(source, "length")
			n := int(vm.jsToNumber(lengthVal).Flt)
			out := make([]Value, 0, n)
			for i := 0; i < n; i++ {
				out = append(out, vm.jsIndexGet(source, NewInteger(int64(i))))
			}
			return out
		case "Set":
			// Set: iterate over unique members stored in jsSetItems.
			if setMap, ok := vm.jsSetItems[source.Num]; ok {
				out := make([]Value, 0, len(setMap))
				for _, v := range setMap {
					out = append(out, v)
				}
				return out
			}
		case "Map":
			// Map: iterate over [key, value] pairs stored in jsMapItems.
			if mapData, ok := vm.jsMapItems[source.Num]; ok {
				out := make([]Value, 0, len(mapData))
				for k, v := range mapData {
					out = append(out, ValueFromVBArray(NewVBArrayFromValues(0, []Value{NewString(k), v})))
				}
				return out
			}
		default:
			// Typed arrays are iterable
			if jsIsTypedArrayType(class) {
				if vals := vm.jsTypedArrayValues(source); vals != nil {
					return vals
				}
			}
		}
	}

	// Fallback: ES6 Iteration Protocol
	itKey := jsSymbolPropertyPrefix + strconv.FormatInt(jsWellKnownSymbolIterator, 10)
	itFn, _ := vm.jsMemberGet(source, itKey)
	if itFn.Type == VTJSFunction || itFn.Type == VTJSObject {
		itObj := vm.jsCall(itFn, source, nil)
		if itObj.Type == VTJSObject {
			out := make([]Value, 0)
			for {
				result, handled := vm.jsCallMember(itObj, "next", nil)
				if !handled || result.Type != VTJSObject {
					break
				}
				doneVal, _ := vm.jsMemberGet(result, "done")
				if vm.jsTruthy(doneVal) {
					break
				}
				val, _ := vm.jsMemberGet(result, "value")
				out = append(out, val)
			}
			return out
		}
	}

	return nil
}

func (vm *VM) jsSuperCall(args []Value) Value {
	if len(vm.jsCallStack) == 0 {
		vm.jsThrowReferenceError("super() call outside of function")
		return Value{Type: VTJSUndefined}
	}
	currentFrameIdx := len(vm.jsCallStack) - 1
	currentFrame := &vm.jsCallStack[currentFrameIdx]
	currentFn := currentFrame.fn
	if currentFn.Type != VTJSFunction {
		vm.jsThrowReferenceError("super() call outside of class constructor")
		return Value{Type: VTJSUndefined}
	}

	// 1. Get the super constructor
	superCtor := vm.jsGetPrototypeValue(currentFn)
	if superCtor.Type != VTJSFunction && superCtor.Type != VTJSObject {
		vm.jsThrowTypeError("Super constructor is not a constructor")
		return Value{Type: VTJSUndefined}
	}

	// 2. Call the super constructor
	// In ES6, super() call in derived constructor uses new.target
	newTarget := vm.jsNewTarget

	// Call the constructor as a super call
	return vm.jsConstruct(superCtor, args, newTarget, true)
}

func (vm *VM) jsCall(callee Value, thisVal Value, args []Value) Value {
	switch callee.Type {
	case VTJSFunction:
		if closure, ok := vm.jsFunctionItems[callee.Num]; ok && closure != nil {
			if closure.isClassConstructor {
				vm.jsThrowTypeError("Class constructor cannot be invoked without 'new'")
				return Value{Type: VTJSUndefined}
			}
			if closure.isGenerator {
				return vm.jsCreateGeneratorObject(callee, thisVal, args)
			}
			if closure.isAsync {
				return vm.jsAsyncCall(callee, thisVal, args)
			}
			if closure.isBound {
				callArgs := closure.boundArgs
				if len(args) > 0 {
					merged := make([]Value, 0, len(callArgs)+len(args))
					merged = append(merged, callArgs...)
					merged = append(merged, args...)
					callArgs = merged
				}
				return vm.jsCall(closure.boundFn, closure.boundThis, callArgs)
			}
			child := vm.cloneForExecuteLocal(len(vm.bytecode))
			if child.jsBeginFunctionCall(callee, thisVal, args, Value{Type: VTJSUndefined}, false, Value{Type: VTJSUndefined}, false) {
				if err := child.Run(); err != nil {
					vm.syncExecuteGlobalState(child)
					if vmErr, ok := err.(*VMError); ok {
						panic(vmErr)
					}
					vm.raise(vbscript.InternalError, err.Error())
					return Value{Type: VTJSUndefined}
				}
				result := Value{Type: VTJSUndefined}
				if child.sp >= 0 {
					result = child.stack[child.sp]
				}
				vm.syncExecuteGlobalState(child)
				return result
			}
			return Value{Type: VTJSUndefined}
		}
		// Native JScript functions (without closures)
		ctorName := vm.jsObjectStringProperty(callee, "__js_ctor")
		switch ctorName {
		case "ArrayValues":
			return vm.jsCreateArrayIterator(thisVal, 0)
		case "StringIteratorFactory":
			return vm.jsCreateStringIterator(vm.valueToString(thisVal))
		case "PromiseResolve":
			promise := vm.jsObjectItems[callee.Num]["__js_promise"]
			resolution := jsArgOrUndefined(args, 0)
			vm.jsResolvePromise(promise, resolution)
			return Value{Type: VTJSUndefined}
		case "PromiseReject":
			promise := vm.jsObjectItems[callee.Num]["__js_promise"]
			reason := jsArgOrUndefined(args, 0)
			vm.jsRejectPromise(promise, reason)
			return Value{Type: VTJSUndefined}
		case "PromiseStaticResolve":
			return vm.jsPromiseStaticResolve(args)
		case "PromiseStaticReject":
			return vm.jsPromiseStaticReject(args)
		case "PromiseStaticAll":
			return vm.jsPromiseStaticAll(args)
		case "PromiseStaticRace":
			return vm.jsPromiseStaticRace(args)
		case "PromiseAllResolver":
			vm.jsHandlePromiseAllResolver(callee, args)
			return Value{Type: VTJSUndefined}
		case "PromiseFinallyHandler":
			return vm.jsHandlePromiseFinallyHandler(callee, args)
		case "PromiseConstantHandler":
			return vm.jsHandlePromiseConstantHandler(callee, args)
		case "PromisePrototypeThen":
			return vm.jsPromiseThen(thisVal, args)
		case "PromisePrototypeCatch":
			return vm.jsPromiseCatch(thisVal, args)
		case "PromisePrototypeFinally":
			return vm.jsPromiseFinally(thisVal, args)
		case "GeneratorPrototypeNext":
			return vm.jsHandleGeneratorNext(thisVal, args)
		case "GeneratorPrototypeThrow":
			return vm.jsHandleGeneratorThrow(thisVal, args)
		case "GeneratorPrototypeReturn":
			return vm.jsHandleGeneratorReturn(thisVal, args)
		}
		return Value{Type: VTJSUndefined}
	case VTNativeObject:
		return vm.dispatchNativeCall(callee.Num, "", args)
	case VTJSObject:
		ctorName := vm.jsObjectStringProperty(callee, "__js_ctor")
		switch ctorName {
		case "String":
			if len(args) == 0 {
				return NewString("")
			}
			return NewString(vm.valueToString(args[0]))
		case "Symbol":
			description := ""
			if len(args) > 0 && args[0].Type != VTJSUndefined {
				description = vm.valueToString(args[0])
			}
			return Value{Type: VTSymbol, Num: vm.jsAllocSymbolID(), Str: description}
		case "Date":
			return NewString(time.Now().In(builtinCurrentLocation(vm)).Format("Mon Jan 02 2006 15:04:05 GMT-0700"))
		case "RegExp":
			if len(args) > 0 {
				return vm.jsNew(callee, args)
			}
			return Value{Type: VTJSUndefined}
		case "Set", "Map":
			vm.jsThrowTypeError(fmt.Sprintf("Constructor %s requires 'new'", ctorName))
			return Value{Type: VTJSUndefined}
		case "ArrayBuffer", "DataView",
			"Int8Array", "Uint8Array", "Uint8ClampedArray",
			"Int16Array", "Uint16Array",
			"Int32Array", "Uint32Array",
			"Float32Array", "Float64Array",
			"BigInt64Array", "BigUint64Array":
			vm.jsThrowTypeError(fmt.Sprintf("Constructor %s requires 'new'", ctorName))
			return Value{Type: VTJSUndefined}
		case "isNaN":
			if len(args) == 0 {
				return NewBool(true)
			}
			return NewBool(math.IsNaN(vm.jsToNumber(args[0]).Flt))
		case "isFinite":
			if len(args) == 0 {
				return NewBool(false)
			}
			num := vm.jsToNumber(args[0]).Flt
			return NewBool(!math.IsNaN(num) && !math.IsInf(num, 0))
		case "parseInt":
			return vm.jsParseIntES5(args)
		case "parseFloat":
			return vm.jsParseFloatES5(args)
		}
		return Value{Type: VTJSUndefined}
	case VTBuiltin:
		idx := int(callee.Num)
		if idx < 0 || idx >= len(BuiltinRegistry) {
			return Value{Type: VTJSUndefined}
		}
		if idx < len(BuiltinNames) && strings.EqualFold(BuiltinNames[idx], "Eval") {
			return vm.jsEval(args)
		}
		result, err := BuiltinRegistry[idx](vm, args)
		if err != nil {
			vm.raise(vbscript.InternalError, err.Error())
			return Value{Type: VTJSUndefined}
		}
		return result
	case VTObject:
		defaultProperty, ok := vm.resolveRuntimeClassPropertyGet(callee, "__default__", len(args), true)
		if ok {
			if vm.beginUserSubCall(defaultProperty, args, false, callee.Num) {
				return Value{Type: VTJSUndefined}
			}
		}
		return Value{Type: VTJSUndefined}
	default:
		return Value{Type: VTJSUndefined}
	}
}

func (vm *VM) jsThrow(v Value) {
	if len(vm.jsTryStack) == 0 {
		vm.raise(vbscript.InternalError, "Unhandled JScript exception")
		vm.push(Value{Type: VTJSUndefined})
		return
	}
	target := vm.jsTryStack[len(vm.jsTryStack)-1]
	vm.jsTryStack = vm.jsTryStack[:len(vm.jsTryStack)-1]
	vm.jsErrStack = append(vm.jsErrStack, v)
	vm.ip = target
}

// jsThrowTypeError throws a JScript TypeError that can be caught by a JS try/catch.
// If no active catch handler exists, raises a VBScript TypeMismatch error instead.
func (vm *VM) jsThrowTypeError(msg string) {
	if len(vm.jsTryStack) == 0 {
		vm.raise(vbscript.TypeMismatch, msg)
		return
	}
	target := vm.jsTryStack[len(vm.jsTryStack)-1]
	vm.jsTryStack = vm.jsTryStack[:len(vm.jsTryStack)-1]
	vm.jsErrStack = append(vm.jsErrStack, NewString("TypeError: "+msg))
	vm.ip = target
}

// jsThrowReferenceError throws a JScript ReferenceError that can be caught by a JS try/catch.
// If no active catch handler exists, raises a VBScript VariableNotDefined error instead.
func (vm *VM) jsThrowReferenceError(msg string) {
	if len(vm.jsTryStack) == 0 {
		vm.raise(vbscript.VariableNotDefined, msg)
		return
	}
	target := vm.jsTryStack[len(vm.jsTryStack)-1]
	vm.jsTryStack = vm.jsTryStack[:len(vm.jsTryStack)-1]
	vm.jsErrStack = append(vm.jsErrStack, NewString("ReferenceError: "+msg))
	vm.ip = target
}

// jsToNumber converts a Value to a numeric value (VTDouble) following JScript semantics.
func (vm *VM) jsToNumber(v Value) Value {
	switch v.Type {
	case VTJSBigInt:
		vm.jsThrowTypeError("Cannot convert a BigInt value to a number")
		return NewDouble(math.NaN())
	case VTJSUndefined:
		return NewDouble(math.NaN())
	case VTNull, VTEmpty:
		return NewDouble(0)
	case VTInteger:
		return Value{Type: VTDouble, Flt: float64(v.Num)}
	case VTDouble:
		return v
	case VTString:
		trimmed := strings.TrimSpace(v.Str)
		if trimmed == "" {
			return NewDouble(0)
		}
		parsed, err := strconv.ParseFloat(trimmed, 64)
		if err != nil {
			return NewDouble(math.NaN())
		}
		return NewDouble(parsed)
	case VTBool:
		if v.Num != 0 {
			return NewDouble(1)
		}
		return NewDouble(0)
	case VTDate:
		return NewDouble(float64(v.Num) / float64(time.Millisecond))
	default:
		return NewDouble(math.NaN())
	}
}

// jsToInt32 converts a value to a 32-bit signed integer for bitwise operations.
func (vm *VM) jsToInt32(v Value) int32 {
	num := vm.jsToNumber(v).Flt
	return int32(num)
}

// jsToUint32 converts a value to a 32-bit unsigned integer for bitwise operations.
func (vm *VM) jsToUint32(v Value) uint32 {
	num := vm.jsToNumber(v).Flt
	return uint32(int32(num))
}

// jsAdd implements JScript '+' operator (string concatenation or numeric addition).
func (vm *VM) jsAdd(a Value, b Value) Value {
	a = resolveCallable(vm, a)
	b = resolveCallable(vm, b)
	if a.Type == VTJSBigInt && b.Type == VTJSBigInt {
		res := new(big.Int).Add(a.Big, b.Big)
		return NewBigInt(res)
	}
	if a.Type == VTJSBigInt || b.Type == VTJSBigInt {
		if a.Type == VTString || b.Type == VTString {
			// Fall through to string concatenation
		} else {
			vm.jsThrowTypeError("Cannot mix BigInt and other types, use explicit conversions")
			return Value{Type: VTJSUndefined}
		}
	}
	return vm.jsAddValues(a, b)
}

// jsSubtract implements JScript '-' operator (numeric subtraction).
func (vm *VM) jsSubtract(a Value, b Value) Value {
	if a.Type == VTJSBigInt && b.Type == VTJSBigInt {
		res := new(big.Int).Sub(a.Big, b.Big)
		return NewBigInt(res)
	}
	if a.Type == VTJSBigInt || b.Type == VTJSBigInt {
		vm.jsThrowTypeError("Cannot mix BigInt and other types, use explicit conversions")
		return Value{Type: VTJSUndefined}
	}
	if a.Type == VTInteger && b.Type == VTInteger {
		if diff, ok := jsSubtractIntegersNoOverflow(a.Num, b.Num); ok {
			return NewInteger(diff)
		}
		return NewDouble(float64(a.Num) - float64(b.Num))
	}
	aNum := vm.jsToNumber(a).Flt
	bNum := vm.jsToNumber(b).Flt
	return NewDouble(aNum - bNum)
}

// jsMultiply implements JScript '*' operator.
func (vm *VM) jsMultiply(a Value, b Value) Value {
	if a.Type == VTJSBigInt && b.Type == VTJSBigInt {
		res := new(big.Int).Mul(a.Big, b.Big)
		return NewBigInt(res)
	}
	if a.Type == VTJSBigInt || b.Type == VTJSBigInt {
		vm.jsThrowTypeError("Cannot mix BigInt and other types, use explicit conversions")
		return Value{Type: VTJSUndefined}
	}
	aNum := vm.jsToNumber(a).Flt
	bNum := vm.jsToNumber(b).Flt
	return NewDouble(aNum * bNum)
}

// jsDivide implements JScript '/' operator.
func (vm *VM) jsDivide(a Value, b Value) Value {
	if a.Type == VTJSBigInt && b.Type == VTJSBigInt {
		if b.Big.Sign() == 0 {
			vm.jsThrowTypeError("Division by zero")
			return Value{Type: VTJSUndefined}
		}
		res := new(big.Int).Div(a.Big, b.Big)
		return NewBigInt(res)
	}
	if a.Type == VTJSBigInt || b.Type == VTJSBigInt {
		vm.jsThrowTypeError("Cannot mix BigInt and other types, use explicit conversions")
		return Value{Type: VTJSUndefined}
	}
	aNum := vm.jsToNumber(a).Flt
	bNum := vm.jsToNumber(b).Flt
	if math.IsNaN(aNum) || math.IsNaN(bNum) {
		return NewDouble(math.NaN())
	}
	if bNum == 0 {
		if aNum == 0 {
			return NewDouble(math.NaN())
		}
		if aNum > 0 {
			return NewDouble(math.Inf(1))
		}
		return NewDouble(math.Inf(-1))
	}
	return NewDouble(aNum / bNum)
}

// jsModulo implements JScript '%' operator.
func (vm *VM) jsModulo(a Value, b Value) Value {
	if a.Type == VTJSBigInt && b.Type == VTJSBigInt {
		if b.Big.Sign() == 0 {
			vm.jsThrowTypeError("Division by zero")
			return Value{Type: VTJSUndefined}
		}
		res := new(big.Int).Mod(a.Big, b.Big)
		return NewBigInt(res)
	}
	if a.Type == VTJSBigInt || b.Type == VTJSBigInt {
		vm.jsThrowTypeError("Cannot mix BigInt and other types, use explicit conversions")
		return Value{Type: VTJSUndefined}
	}
	aNum := vm.jsToNumber(a).Flt
	bNum := vm.jsToNumber(b).Flt
	if bNum == 0 || math.IsNaN(aNum) || math.IsNaN(bNum) || math.IsInf(aNum, 0) || math.IsInf(bNum, 0) {
		return NewDouble(math.NaN()) // NaN for modulo by zero
	}
	return NewDouble(math.Mod(aNum, bNum))
}

// jsExponent implements JScript '**' operator.
func (vm *VM) jsExponent(a Value, b Value) Value {
	if a.Type == VTJSBigInt && b.Type == VTJSBigInt {
		if b.Big.Sign() < 0 {
			vm.jsThrowTypeError("Exponent must be positive")
			return Value{Type: VTJSUndefined}
		}
		if !b.Big.IsUint64() {
			vm.jsThrowTypeError("Exponent too large")
			return Value{Type: VTJSUndefined}
		}
		res := new(big.Int).Exp(a.Big, b.Big, nil)
		return NewBigInt(res)
	}
	if a.Type == VTJSBigInt || b.Type == VTJSBigInt {
		vm.jsThrowTypeError("Cannot mix BigInt and other types, use explicit conversions")
		return Value{Type: VTJSUndefined}
	}
	aNum := vm.jsToNumber(a).Flt
	bNum := vm.jsToNumber(b).Flt
	return NewDouble(math.Pow(aNum, bNum))
}

// jsNegate implements JScript unary '-' operator.
func (vm *VM) jsNegate(v Value) Value {
	if v.Type == VTJSBigInt {
		res := new(big.Int).Neg(v.Big)
		return NewBigInt(res)
	}
	num := vm.jsToNumber(v).Flt
	return NewDouble(-num)
}

// jsBitwiseAnd implements JScript '&' operator.
func (vm *VM) jsBitwiseAnd(a Value, b Value) Value {
	aInt := vm.jsToInt32(a)
	bInt := vm.jsToInt32(b)
	return NewInteger(int64(aInt & bInt))
}

// jsBitwiseOr implements JScript '|' operator.
func (vm *VM) jsBitwiseOr(a Value, b Value) Value {
	aInt := vm.jsToInt32(a)
	bInt := vm.jsToInt32(b)
	return NewInteger(int64(aInt | bInt))
}

// jsBitwiseXor implements JScript '^' operator.
func (vm *VM) jsBitwiseXor(a Value, b Value) Value {
	aInt := vm.jsToInt32(a)
	bInt := vm.jsToInt32(b)
	return NewInteger(int64(aInt ^ bInt))
}

// jsBitwiseNot implements JScript '~' operator (bitwise NOT).
func (vm *VM) jsBitwiseNot(v Value) Value {
	vInt := vm.jsToInt32(v)
	return NewInteger(int64(^vInt))
}

// jsLeftShift implements JScript '<<' operator.
func (vm *VM) jsLeftShift(a Value, b Value) Value {
	aInt := vm.jsToInt32(a)
	bInt := vm.jsToUint32(b) & 0x1f // Only use lower 5 bits
	return NewInteger(int64(aInt << bInt))
}

// jsRightShift implements JScript '>>' operator (sign-extending right shift).
func (vm *VM) jsRightShift(a Value, b Value) Value {
	aInt := vm.jsToInt32(a)
	bInt := vm.jsToUint32(b) & 0x1f // Only use lower 5 bits
	return NewInteger(int64(aInt >> bInt))
}

// jsUnsignedRightShift implements JScript '>>>' operator (zero-filling right shift).
func (vm *VM) jsUnsignedRightShift(a Value, b Value) Value {
	aUint := vm.jsToUint32(a)
	bInt := vm.jsToUint32(b) & 0x1f // Only use lower 5 bits
	return NewInteger(int64(aUint >> bInt))
}

// jsLess implements JScript '<' operator.
func (vm *VM) jsLess(a Value, b Value) Value {
	// Fast path: both integers — avoid jsToNumber call and int→float conversion.
	if a.Type == VTInteger && b.Type == VTInteger {
		return NewBool(a.Num < b.Num)
	}
	// Fast path: both doubles.
	if a.Type == VTDouble && b.Type == VTDouble {
		return NewBool(a.Flt < b.Flt)
	}
	// Fast path: integer vs double mixed.
	if a.Type == VTInteger && b.Type == VTDouble {
		return NewBool(float64(a.Num) < b.Flt)
	}
	if a.Type == VTDouble && b.Type == VTInteger {
		return NewBool(a.Flt < float64(b.Num))
	}
	aNum := vm.jsToNumber(a).Flt
	bNum := vm.jsToNumber(b).Flt
	return NewBool(aNum < bNum)
}

// jsGreater implements JScript '>' operator.
func (vm *VM) jsGreater(a Value, b Value) Value {
	// Fast path: both integers.
	if a.Type == VTInteger && b.Type == VTInteger {
		return NewBool(a.Num > b.Num)
	}
	// Fast path: both doubles.
	if a.Type == VTDouble && b.Type == VTDouble {
		return NewBool(a.Flt > b.Flt)
	}
	// Fast path: integer vs double mixed.
	if a.Type == VTInteger && b.Type == VTDouble {
		return NewBool(float64(a.Num) > b.Flt)
	}
	if a.Type == VTDouble && b.Type == VTInteger {
		return NewBool(a.Flt > float64(b.Num))
	}
	aNum := vm.jsToNumber(a).Flt
	bNum := vm.jsToNumber(b).Flt
	return NewBool(aNum > bNum)
}

// jsLessEqual implements JScript '<=' operator.
func (vm *VM) jsLessEqual(a Value, b Value) Value {
	// Fast path: both integers.
	if a.Type == VTInteger && b.Type == VTInteger {
		return NewBool(a.Num <= b.Num)
	}
	// Fast path: both doubles.
	if a.Type == VTDouble && b.Type == VTDouble {
		return NewBool(a.Flt <= b.Flt)
	}
	// Fast path: integer vs double mixed.
	if a.Type == VTInteger && b.Type == VTDouble {
		return NewBool(float64(a.Num) <= b.Flt)
	}
	if a.Type == VTDouble && b.Type == VTInteger {
		return NewBool(a.Flt <= float64(b.Num))
	}
	aNum := vm.jsToNumber(a).Flt
	bNum := vm.jsToNumber(b).Flt
	return NewBool(aNum <= bNum)
}

// jsGreaterEqual implements JScript '>=' operator.
func (vm *VM) jsGreaterEqual(a Value, b Value) Value {
	// Fast path: both integers.
	if a.Type == VTInteger && b.Type == VTInteger {
		return NewBool(a.Num >= b.Num)
	}
	// Fast path: both doubles.
	if a.Type == VTDouble && b.Type == VTDouble {
		return NewBool(a.Flt >= b.Flt)
	}
	// Fast path: integer vs double mixed.
	if a.Type == VTInteger && b.Type == VTDouble {
		return NewBool(float64(a.Num) >= b.Flt)
	}
	if a.Type == VTDouble && b.Type == VTInteger {
		return NewBool(a.Flt >= float64(b.Num))
	}
	aNum := vm.jsToNumber(a).Flt
	bNum := vm.jsToNumber(b).Flt
	return NewBool(aNum >= bNum)
}

// jsLooseEqual implements JScript '==' (loose equality) operator.
func (vm *VM) jsLooseEqual(a Value, b Value) Value {
	a = resolveCallable(vm, a)
	b = resolveCallable(vm, b)

	if vm.jsStrictEquals(a, b) {
		return NewBool(true)
	}

	aNullish := a.Type == VTNull || a.Type == VTJSUndefined
	bNullish := b.Type == VTNull || b.Type == VTJSUndefined
	if aNullish || bNullish {
		return NewBool(aNullish && bNullish)
	}

	if a.Type == VTBool {
		a = vm.jsToNumber(a)
	}
	if b.Type == VTBool {
		b = vm.jsToNumber(b)
	}

	aNum := vm.jsToNumber(a).Flt
	bNum := vm.jsToNumber(b).Flt
	if math.IsNaN(aNum) || math.IsNaN(bNum) {
		return NewBool(false)
	}
	return NewBool(aNum == bNum)
}

// jsLooseNotEqual implements JScript '!=' (loose inequality) operator.
func (vm *VM) jsLooseNotEqual(a Value, b Value) Value {
	eq := vm.jsLooseEqual(a, b)
	return NewBool(eq.Num == 0)
}

// jsLogicalAnd implements JScript '&&' operator.
func (vm *VM) jsLogicalAnd(a Value, b Value) Value {
	if !vm.jsTruthy(a) {
		return a
	}
	return b
}

// jsLogicalOr implements JScript '||' operator.
func (vm *VM) jsLogicalOr(a Value, b Value) Value {
	if vm.jsTruthy(a) {
		return a
	}
	return b
}

// jsLogicalNot implements JScript '!' operator.
func (vm *VM) jsLogicalNot(v Value) Value {
	return NewBool(!vm.jsTruthy(v))
}

// jsCoalesce implements JScript '??' operator.
func (vm *VM) jsCoalesce(a Value, b Value) Value {
	if a.Type == VTNull || a.Type == VTJSUndefined {
		return b
	}
	return a
}

// jsMemberIndexGet implements member[index] access.
func (vm *VM) jsMemberIndexGet(obj Value, index Value, memberName string) Value {
	container, deferred := vm.jsMemberGet(obj, memberName)
	if deferred {
		return Value{Type: VTJSUndefined}
	}
	return vm.jsIndexGet(container, index)
}

// jsMemberIndexSet implements member[index] = value assignment.
func (vm *VM) jsMemberIndexSet(obj Value, index Value, value Value, memberName string) {
	container, deferred := vm.jsMemberGet(obj, memberName)
	if deferred {
		return
	}
	vm.jsIndexSet(container, index, value)
}

func (vm *VM) jsUpdateMember(target Value, member string, delta float64, post bool) Value {
	current, deferred := vm.jsMemberGet(target, member)
	if deferred {
		return Value{Type: VTJSUndefined}
	}
	next := vm.jsToNumber(current)
	next.Flt += delta
	vm.jsMemberSet(target, member, next)
	if post {
		return current
	}
	return next
}

func (vm *VM) jsUpdateIndex(target Value, index Value, delta float64, post bool) Value {
	current := vm.jsIndexGet(target, index)
	next := vm.jsToNumber(current)
	next.Flt += delta
	vm.jsIndexSet(target, index, next)
	if post {
		return current
	}
	return next
}

func (vm *VM) jsNew(constructor Value, args []Value) Value {
	return vm.jsConstruct(constructor, args, constructor, false)
}

func (vm *VM) jsConstruct(constructor Value, args []Value, newTarget Value, isSuper bool) Value {
	if constructor.Type == VTJSFunction {
		closure := vm.jsFunctionItems[constructor.Num]
		if closure != nil && closure.isDerived {
			// Derived class constructor: 'this' is uninitialized until super() is called.
			if vm.jsBeginFunctionCall(constructor, Value{Type: VTJSUninitialized}, args, Value{Type: VTJSUninitialized}, true, newTarget, isSuper) {
				return Value{Type: VTJSUndefined}
			}
			return Value{Type: VTJSUninitialized}
		}

		instanceID := vm.allocJSID()
		vm.jsObjectItems[instanceID] = make(map[string]Value, 8)
		vm.jsPropertyItems[instanceID] = make(map[string]jsPropertyDescriptor, 8)
		instance := Value{Type: VTJSObject, Num: instanceID}
		// Use prototype from newTarget (which might be the subclass if this is a base constructor call via super())
		proto, deferred := vm.jsMemberGet(newTarget, "prototype")
		if !deferred && proto.Type == VTJSObject {
			vm.jsObjectItems[instanceID]["__js_proto"] = proto
		} else {
			fallback := vm.jsGetIntrinsicPrototype("Object")
			if fallback.Type == VTJSObject {
				vm.jsObjectItems[instanceID]["__js_proto"] = fallback
			}
		}
		if vm.jsBeginFunctionCall(constructor, instance, args, instance, true, newTarget, isSuper) {
			return Value{Type: VTJSUndefined}
		}
		return instance
	}
	if constructor.Type == VTJSObject {
		ctorName := vm.jsObjectStringProperty(constructor, "__js_ctor")
		switch ctorName {
		case "Symbol":
			vm.jsThrowTypeError("Symbol is not a constructor")
			return Value{Type: VTJSUndefined}
		case "Date":
			if len(args) == 0 {
				return NewDate(time.Now().In(builtinCurrentLocation(vm)))
			}
			year := int(vm.jsToNumber(args[0]).Flt)
			month := 0
			day := 1
			hour := 0
			minute := 0
			second := 0
			millisecond := 0
			if len(args) > 1 {
				month = int(vm.jsToNumber(args[1]).Flt)
			}
			if len(args) > 2 {
				day = int(vm.jsToNumber(args[2]).Flt)
			}
			if len(args) > 3 {
				hour = int(vm.jsToNumber(args[3]).Flt)
			}
			if len(args) > 4 {
				minute = int(vm.jsToNumber(args[4]).Flt)
			}
			if len(args) > 5 {
				second = int(vm.jsToNumber(args[5]).Flt)
			}
			if len(args) > 6 {
				millisecond = int(vm.jsToNumber(args[6]).Flt)
			}
			loc := builtinCurrentLocation(vm)
			t := time.Date(year, time.Month(month+1), day, hour, minute, second, millisecond*int(time.Millisecond), loc)
			return NewDate(t)
		case "RegExp":
			pattern := ""
			flags := ""
			if len(args) > 0 {
				pattern = vm.valueToString(args[0])
			}
			if len(args) > 1 {
				flags = vm.valueToString(args[1])
			}
			objID := vm.allocJSID()
			obj := make(map[string]Value, 3)
			obj["__js_type"] = NewString("RegExp")
			obj["pattern"] = NewString(pattern)
			obj["flags"] = NewString(flags)
			vm.jsObjectItems[objID] = obj
			vm.jsPropertyItems[objID] = make(map[string]jsPropertyDescriptor, 4)
			return Value{Type: VTJSObject, Num: objID}
		case "Enumerator":
			source := Value{Type: VTJSUndefined}
			if len(args) > 0 {
				source = args[0]
			}
			objID := vm.allocJSID()
			obj := make(map[string]Value, 4)
			obj["__js_type"] = NewString("Enumerator")
			obj["__js_enum_source"] = source
			obj["__js_enum_index"] = NewInteger(0)
			if source.Type == VTJSObject {
				keys := vm.jsEnumerateForInKeys(source)
				keyVals := make([]Value, len(keys))
				for i := 0; i < len(keys); i++ {
					keyVals[i] = NewString(keys[i])
				}
				obj["__js_enum_keys"] = ValueFromVBArray(NewVBArrayFromValues(0, keyVals))
			}
			vm.jsObjectItems[objID] = obj
			vm.jsPropertyItems[objID] = make(map[string]jsPropertyDescriptor, 6)
			return Value{Type: VTJSObject, Num: objID}
		case "VBArray":
			source := Value{Type: VTJSUndefined}
			if len(args) > 0 {
				source = args[0]
			}
			objID := vm.allocJSID()
			obj := make(map[string]Value, 2)
			obj["__js_type"] = NewString("VBArray")
			obj["__js_vbarray_source"] = source
			vm.jsObjectItems[objID] = obj
			vm.jsPropertyItems[objID] = make(map[string]jsPropertyDescriptor, 3)
			if proto := vm.jsGetIntrinsicPrototype("Array"); proto.Type == VTJSObject {
				obj["__js_proto"] = proto
			}
			return Value{Type: VTJSObject, Num: objID}
		case "Set":
			objID := vm.allocJSID()
			obj := make(map[string]Value, 2)
			obj["__js_type"] = NewString("Set")
			obj["__js_ctor"] = NewString("Set")
			if proto := vm.jsGetIntrinsicPrototype("Set"); proto.Type == VTJSObject {
				obj["__js_proto"] = proto
			}
			vm.jsObjectItems[objID] = obj
			vm.jsPropertyItems[objID] = make(map[string]jsPropertyDescriptor, 4)
			vm.jsSetItems[objID] = make(map[string]Value, 8)
			setObj := Value{Type: VTJSObject, Num: objID}
			if len(args) > 0 && args[0].Type != VTJSUndefined && args[0].Type != VTNull {
				length, hasLength, deferred := vm.jsArrayLikeLength(args[0])
				if deferred {
					return Value{Type: VTJSUndefined}
				}
				if hasLength {
					for i := 0; i < length; i++ {
						if v, ok := vm.jsArrayLikeGetIndex(args[0], i); ok {
							vm.jsSetItems[objID][vm.jsValueMapKey(v)] = v
						}
					}
				}
			}
			return setObj
		case "Map":
			objID := vm.allocJSID()
			obj := make(map[string]Value, 2)
			obj["__js_type"] = NewString("Map")
			obj["__js_ctor"] = NewString("Map")
			if proto := vm.jsGetIntrinsicPrototype("Map"); proto.Type == VTJSObject {
				obj["__js_proto"] = proto
			}
			vm.jsObjectItems[objID] = obj
			vm.jsPropertyItems[objID] = make(map[string]jsPropertyDescriptor, 4)
			vm.jsMapItems[objID] = make(map[string]Value, 8)
			mapObj := Value{Type: VTJSObject, Num: objID}
			if len(args) > 0 && args[0].Type != VTJSUndefined && args[0].Type != VTNull {
				length, hasLength, deferred := vm.jsArrayLikeLength(args[0])
				if deferred {
					return Value{Type: VTJSUndefined}
				}
				if hasLength {
					for i := 0; i < length; i++ {
						entry, ok := vm.jsArrayLikeGetIndex(args[0], i)
						if !ok || entry.Type != VTArray || entry.Arr == nil || len(entry.Arr.Values) < 2 {
							continue
						}
						vm.jsMapItems[objID][vm.jsValueMapKey(entry.Arr.Values[0])] = entry.Arr.Values[1]
					}
				}
			}
			return mapObj
		case "Promise":
			return vm.jsNewPromise(args)
		case "ArrayBuffer":
			byteLength := 0
			if len(args) > 0 {
				byteLength = int(vm.jsToNumber(args[0]).Flt)
			}
			return vm.jsNewArrayBuffer(byteLength)
		case "DataView":
			return vm.jsNewDataView(args)
		case "Int8Array", "Uint8Array", "Uint8ClampedArray",
			"Int16Array", "Uint16Array",
			"Int32Array", "Uint32Array",
			"Float32Array", "Float64Array",
			"BigInt64Array", "BigUint64Array":
			return vm.jsNewTypedArray(ctorName, args)
		}
	}

	objID := vm.allocJSID()
	vm.jsObjectItems[objID] = make(map[string]Value, 8)
	vm.jsPropertyItems[objID] = make(map[string]jsPropertyDescriptor, 8)
	return Value{Type: VTJSObject, Num: objID}
}

// jsMemberDelete implements the 'delete' operator for object properties.
func (vm *VM) jsMemberDelete(obj Value, member string) bool {
	if obj.Type == VTJSObject {
		desc, hasDesc := vm.jsGetDescriptor(obj.Num, member)
		if hasDesc && !desc.Configurable {
			return false
		}
		if jsObj, ok := vm.jsObjectItems[obj.Num]; ok {
			delete(jsObj, member)
			if props, ok := vm.jsPropertyItems[obj.Num]; ok {
				delete(props, member)
			}
			return true
		}
	}
	return false
}

// jsIndexGet implements array[index] access.
func (vm *VM) jsIndexGet(arr Value, index Value) Value {
	if index.Type == VTSymbol {
		key := vm.jsPropertyKeyFromValue(index)
		val, _ := vm.jsMemberGet(arr, key)
		return val
	}
	switch arr.Type {
	case VTArray:
		if arr.Arr == nil {
			return Value{Type: VTJSUndefined}
		}
		indexNum := int(vm.jsToNumber(index).Flt)
		adjustedIndex := indexNum - arr.Arr.Lower
		if adjustedIndex < 0 || adjustedIndex >= len(arr.Arr.Values) {
			// Fallback to property/prototype lookup (e.g. Symbol.iterator)
			key := vm.jsPropertyKeyFromValue(index)
			val, _ := vm.jsMemberGet(arr, key)
			return val
		}
		return arr.Arr.Values[adjustedIndex]
	case VTJSObject:
		// First try typed array index access.
		if idxInt := int(vm.jsToNumber(index).Flt); true {
			if v, handled := vm.jsTypedArrayIndexGet(arr, idxInt); handled {
				return v
			}
		}
		key := vm.jsPropertyKeyFromValue(index)
		if v, _ := vm.jsMemberGet(arr, key); v.Type != VTJSUndefined {
			return v
		}
		return Value{Type: VTJSUndefined}
	case VTString:
		runes := []rune(arr.Str)
		idx := int(vm.jsToNumber(index).Flt)
		if idx < 0 || idx >= len(runes) {
			// Fallback to property/prototype lookup
			key := vm.jsPropertyKeyFromValue(index)
			val, _ := vm.jsMemberGet(arr, key)
			return val
		}
		ch := string(runes[idx])
		if !vm.jsEnsureStringSize(len(ch)) || !vm.jsChargeStringWork(len(ch)) {
			return Value{Type: VTJSUndefined}
		}
		return NewString(ch)
	default:
		return Value{Type: VTJSUndefined}
	}
}

// jsIndexSet implements array[index] = value assignment.
func (vm *VM) jsIndexSet(arr Value, index Value, value Value) {
	switch arr.Type {
	case VTArray:
		if arr.Arr == nil {
			return
		}
		indexNum := int(vm.jsToNumber(index).Flt)
		adjustedIndex := indexNum - arr.Arr.Lower
		if adjustedIndex >= 0 && adjustedIndex < len(arr.Arr.Values) {
			arr.Arr.Values[adjustedIndex] = value
		}
	case VTJSObject:
		// Try typed array index set first.
		idxInt := int(vm.jsToNumber(index).Flt)
		if vm.jsTypedArrayIndexSet(arr, idxInt, value) {
			return
		}
		key := vm.jsPropertyKeyFromValue(index)
		vm.jsMemberSet(arr, key, value)
	}
}

func (vm *VM) jsCompileRegExp(pattern string, flags string) (*regexp.Regexp, error) {
	flagsLower := strings.ToLower(flags)
	isUnicode := strings.Contains(flagsLower, "u")
	transformed, err := jsparser.TransformRegExp(pattern, false, isUnicode)
	if err != nil {
		// Fallback to literal interpretation if syntax is not understood
		transformed = pattern
	}
	if strings.Contains(flagsLower, "i") {
		transformed = "(?i)" + transformed
	}
	return regexp.Compile(transformed)
}

func jsStringLength(s string) int64 {
	length := int64(0)
	for _, r := range s {
		if r >= 0x10000 {
			length += 2
		} else {
			length += 1
		}
	}
	return length
}

func (vm *VM) jsSuperGet(member string) Value {
	if len(vm.jsCallStack) == 0 {
		return Value{Type: VTJSUndefined}
	}
	currentFn := vm.jsCallStack[len(vm.jsCallStack)-1].fn
	if currentFn.Type != VTJSFunction {
		return Value{Type: VTJSUndefined}
	}
	closure := vm.jsFunctionItems[currentFn.Num]
	if closure == nil || closure.homeObjID == 0 {
		// Not inside a class method
		vm.jsThrowTypeError("super member access outside of class method")
		return Value{Type: VTJSUndefined}
	}

	homeObj := Value{Type: VTJSObject, Num: closure.homeObjID}
	superProto := vm.jsGetPrototypeValue(homeObj)
	if superProto.Type == VTJSUndefined {
		return Value{Type: VTJSUndefined}
	}

	val, deferred := vm.jsMemberGet(superProto, member)
	if deferred {
		// Getter call was initiated
		return Value{Type: VTJSUndefined}
	}
	return val
}

func (vm *VM) jsSuperSet(member string, val Value) {
	if len(vm.jsCallStack) == 0 {
		return
	}
	currentFn := vm.jsCallStack[len(vm.jsCallStack)-1].fn
	if currentFn.Type != VTJSFunction {
		return
	}
	closure := vm.jsFunctionItems[currentFn.Num]
	if closure == nil || closure.homeObjID == 0 {
		vm.jsThrowTypeError("super member assignment outside of class method")
		return
	}

	homeObj := Value{Type: VTJSObject, Num: closure.homeObjID}
	superProto := vm.jsGetPrototypeValue(homeObj)
	if superProto.Type == VTJSUndefined {
		return
	}

	// Super assignment receiver is 'this'
	vm.jsMemberSet(vm.jsThisValue, member, val)
}

func (vm *VM) jsSuperCallMember(member string, args []Value) Value {
	if len(vm.jsCallStack) == 0 {
		return Value{Type: VTJSUndefined}
	}
	currentFn := vm.jsCallStack[len(vm.jsCallStack)-1].fn
	if currentFn.Type != VTJSFunction {
		return Value{Type: VTJSUndefined}
	}
	closure := vm.jsFunctionItems[currentFn.Num]
	if closure == nil || closure.homeObjID == 0 {
		vm.jsThrowTypeError("super member call outside of class method")
		return Value{Type: VTJSUndefined}
	}

	homeObj := Value{Type: VTJSObject, Num: closure.homeObjID}
	superProto := vm.jsGetPrototypeValue(homeObj)
	if superProto.Type == VTJSUndefined {
		vm.jsThrowTypeError("Super prototype is undefined")
		return Value{Type: VTJSUndefined}
	}

	// Resolve the method from superProto
	method, deferred := vm.jsMemberGet(superProto, member)
	if deferred {
		return Value{Type: VTJSUndefined}
	}

	if method.Type != VTJSFunction && method.Type != VTNativeObject && method.Type != VTBuiltin {
		vm.jsThrowTypeError(fmt.Sprintf("super.%s is not a function", member))
		return Value{Type: VTJSUndefined}
	}

	// Call the method with 'this' set to current instance
	return vm.jsCall(method, vm.jsThisValue, args)
}

func (vm *VM) jsSuperIndexGet(index Value) Value {
	return vm.jsSuperGet(vm.jsPropertyKeyFromValue(index))
}

func (vm *VM) jsSuperIndexSet(index Value, val Value) {
	vm.jsSuperSet(vm.jsPropertyKeyFromValue(index), val)
}
