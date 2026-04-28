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
	"encoding/json"
	"fmt"
	"math"
	"math/rand"
	"regexp"
	"sort"
	"strconv"
	"strings"
	"time"

	"g3pix.com.br/axonasp/vbscript"
)

const jsMaxStringBytes = 8 * 1024 * 1024
const jsMaxStringWorkBytes = 2 * 1024 * 1024
const jsInternalPropPrefix = "__js_"
const jsAccessorGetterPrefix = "__js_getter__"
const jsAccessorSetterPrefix = "__js_setter__"

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

type jsFunctionObject struct {
	name      string
	params    []string
	startIP   int
	endIP     int
	envID     int64
	protoID   int64
	isBound   bool
	boundFn   Value
	boundThis Value
	boundArgs []Value
}

type jsCallFrame struct {
	returnIP int
	envID    int64
	thisVal  Value
	tryDepth int
	savedSP  int
	isCtor   bool
	ctorObj  Value
}

type jsForInEnumerator struct {
	keys  []string
	index int
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
			keys[key] = struct{}{}
		}
	}
	if props, ok := vm.jsPropertyItems[target.Num]; ok {
		for key := range props {
			if strings.HasPrefix(key, jsInternalPropPrefix) {
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
	case "Array", "Object", "String", "Date", "RegExp", "Enumerator", "VBArray":
		return true
	default:
		return false
	}
}

func (vm *VM) jsGetPrototypeValue(v Value) Value {
	switch v.Type {
	case VTJSObject:
		if obj, ok := vm.jsObjectItems[v.Num]; ok {
			if proto, exists := obj["__js_proto"]; exists {
				return proto
			}
		}
	case VTJSFunction:
		if fn, ok := vm.jsFunctionItems[v.Num]; ok && fn != nil && fn.protoID != 0 {
			return Value{Type: VTJSObject, Num: fn.protoID}
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
	if !exists || proto.Type != VTJSObject {
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

func (vm *VM) jsJSONStringify(v Value) string {
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
			b.WriteString(vm.jsJSONStringify(v.Arr.Values[i]))
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
			b.WriteString(vm.jsJSONStringify(obj[k]))
		}
		b.WriteByte('}')
		return b.String()
	default:
		encoded, _ := json.Marshal(vm.valueToString(v))
		return string(encoded)
	}
}

func (vm *VM) ensureJSRootEnv() {
	if vm.jsActiveEnvID != 0 {
		return
	}
	rootID := vm.allocJSID()
	bindings := make(map[string]Value, 16)
	bindings["Math"] = vm.jsCreateMathObject()
	bindings["Date"] = vm.jsCreateIntrinsicObject("", "Date")
	bindings["RegExp"] = vm.jsCreateIntrinsicObject("", "RegExp")
	bindings["Enumerator"] = vm.jsCreateIntrinsicObject("", "Enumerator")
	bindings["VBArray"] = vm.jsCreateIntrinsicObject("", "VBArray")
	bindings["String"] = vm.jsCreateIntrinsicObject("", "String")
	bindings["Array"] = vm.jsCreateIntrinsicObject("", "Array")
	bindings["Object"] = vm.jsCreateIntrinsicObject("", "Object")
	bindings["JSON"] = vm.jsCreateIntrinsicObject("", "JSON")
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
	if obj.Type != VTJSObject {
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

func (vm *VM) jsSetName(name string, val Value) {
	vm.ensureJSRootEnv()
	for envID := vm.jsActiveEnvID; envID != 0; {
		env := vm.jsEnvItems[envID]
		if env == nil {
			break
		}
		if _, ok := env.bindings[name]; ok {
			env.bindings[name] = val
			return
		}
		envID = env.parentID
	}
	if idx, ok := vm.lookupJSGlobalIndex(name); ok {
		vm.Globals[idx] = val
		return
	}
	root := vm.jsEnvItems[vm.jsActiveEnvID]
	if root != nil {
		root.bindings[name] = val
	}
}

func (vm *VM) jsGetName(name string) Value {
	vm.ensureJSRootEnv()
	if strings.EqualFold(name, "this") {
		return vm.jsThisValue
	}
	if strings.EqualFold(name, "eval") {
		if idx, ok := GetBuiltinIndex("Eval"); ok {
			return Value{Type: VTBuiltin, Num: int64(idx)}
		}
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
	case VTBool, VTInteger, VTDate, VTNativeObject, VTJSObject, VTJSFunction:
		return a.Num == b.Num
	case VTDouble:
		return a.Flt == b.Flt
	case VTString:
		return a.Str == b.Str
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
	case VTJSFunction:
		return "function"
	case VTJSObject, VTNativeObject, VTObject, VTArray:
		return "object"
	default:
		return "undefined"
	}
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
	if template.Type != VTJSFunctionTemplate {
		return Value{Type: VTJSUndefined}
	}
	id := vm.allocJSID()
	fnVal := Value{Type: VTJSFunction, Num: id}
	proto := vm.jsCreatePrototypeObject(fnVal)
	vm.jsFunctionItems[id] = &jsFunctionObject{
		name:    template.Str,
		params:  append([]string(nil), template.Names...),
		startIP: int(template.Num),
		endIP:   int(template.Flt),
		envID:   vm.jsActiveEnvID,
		protoID: proto.Num,
	}
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

func (vm *VM) jsBeginFunctionCall(fn Value, thisVal Value, args []Value, ctorObj Value, isCtor bool) bool {
	closure, ok := vm.jsFunctionItems[fn.Num]
	if !ok || closure == nil {
		return false
	}
	frame := jsCallFrame{
		returnIP: vm.ip,
		envID:    vm.jsActiveEnvID,
		thisVal:  vm.jsThisValue,
		tryDepth: len(vm.jsTryStack),
		savedSP:  vm.sp,
		isCtor:   isCtor,
		ctorObj:  ctorObj,
	}
	vm.jsCallStack = append(vm.jsCallStack, frame)
	envID := vm.allocJSID()
	bindings := make(map[string]Value, len(closure.params)+1)
	for i := 0; i < len(closure.params); i++ {
		if i < len(args) {
			bindings[closure.params[i]] = args[i]
		} else {
			bindings[closure.params[i]] = Value{Type: VTJSUndefined}
		}
	}
	if _, hasArguments := bindings["arguments"]; !hasArguments {
		argumentsObject := vm.jsCreateArgumentsObject(args)
		bindings["arguments"] = argumentsObject
	}
	vm.jsEnvItems[envID] = &jsEnvFrame{parentID: closure.envID, bindings: bindings}
	vm.jsActiveEnvID = envID
	vm.jsThisValue = thisVal
	vm.ip = closure.startIP
	return true
}

func (vm *VM) jsCreateArgumentsObject(args []Value) Value {
	objID := vm.allocJSID()
	obj := make(map[string]Value, len(args)+1)
	for i := 0; i < len(args); i++ {
		obj[strconv.Itoa(i)] = args[i]
	}
	obj["length"] = NewInteger(int64(len(args)))
	vm.jsObjectItems[objID] = obj
	vm.jsPropertyItems[objID] = make(map[string]jsPropertyDescriptor, len(args)+1)
	return Value{Type: VTJSObject, Num: objID}
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
	vm.ip = frame.returnIP
	vm.sp = frame.savedSP
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
	switch target.Type {
	case VTNativeObject:
		return vm.dispatchMemberGet(target, member), false
	case VTString:
		if strings.EqualFold(member, "length") {
			return NewInteger(int64(len(target.Str))), false
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
	case VTJSFunction:
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

	switch target.Type {
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
			rePattern := pattern
			if strings.Contains(strings.ToLower(flags), "i") {
				rePattern = "(?i)" + rePattern
			}
			re, err := regexp.Compile(rePattern)
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
			rePattern := pattern
			if strings.Contains(strings.ToLower(flags), "i") {
				rePattern = "(?i)" + rePattern
			}
			re, err := regexp.Compile(rePattern)
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
			replacement := ""
			if len(args) > 1 {
				replacement = vm.valueToString(args[1])
			}
			return vm.jsStringReplace(text, args[0], replacement, false), true
		case strings.EqualFold(member, "replaceAll"):
			if len(args) == 0 {
				return NewString(text), true
			}
			replacement := ""
			if len(args) > 1 {
				replacement = vm.valueToString(args[1])
			}
			return vm.jsStringReplace(text, args[0], replacement, true), true
		case strings.EqualFold(member, "trim"):
			trimmed := strings.TrimSpace(text)
			if !vm.jsEnsureStringSize(len(trimmed)) || !vm.jsChargeStringWork(len(trimmed)) {
				return Value{Type: VTJSUndefined}, true
			}
			return NewString(trimmed), true
		}
	case VTArray:
		if target.Arr == nil {
			return Value{Type: VTJSUndefined}, true
		}
		switch {
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
	case VTJSObject:
		switch {
		case strings.EqualFold(member, "hasOwnProperty"):
			return NewBool(vm.jsObjectHasOwnProperty(target, vm.valueToString(jsArgOrUndefined(args, 0)))), true
		case strings.EqualFold(member, "propertyIsEnumerable"):
			return NewBool(vm.jsObjectPropertyIsEnumerable(target, vm.valueToString(jsArgOrUndefined(args, 0)))), true
		case strings.EqualFold(member, "isPrototypeOf"):
			return NewBool(vm.jsObjectIsPrototypeOf(target, jsArgOrUndefined(args, 0))), true
		case strings.EqualFold(member, "toString"):
			return NewString("[object Object]"), true
		case strings.EqualFold(member, "toLocaleString"):
			return NewString("[object Object]"), true
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
			}
		case "Object":
			switch {
			case strings.EqualFold(member, "create"):
				objID := vm.allocJSID()
				obj := make(map[string]Value, 8)
				vm.jsObjectItems[objID] = obj
				vm.jsPropertyItems[objID] = make(map[string]jsPropertyDescriptor, 8)
				if len(args) > 0 {
					obj["__js_proto"] = args[0]
				}
				return Value{Type: VTJSObject, Num: objID}, true
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
			case strings.EqualFold(member, "keys"):
				if len(args) == 0 || (args[0].Type != VTJSObject && args[0].Type != VTJSFunction) {
					return ValueFromVBArray(NewVBArrayFromValues(0, nil)), true
				}
				keys := vm.jsObjectOwnEnumerableKeys(args[0].Num)
				values := make([]Value, len(keys))
				for i := 0; i < len(keys); i++ {
					values[i] = NewString(keys[i])
				}
				return ValueFromVBArray(NewVBArrayFromValues(0, values)), true
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
				if !vm.jsObjectIsExtensible(args[0]) {
					if _, exists := vm.jsGetDescriptor(objID, name); !exists {
						return args[0], true
					}
				}
				descVal := args[2]
				desc := jsPropertyDescriptor{Enumerable: false, Configurable: false, Writable: false}
				if descVal.Type == VTJSObject {
					if dObj, ok := vm.jsObjectItems[descVal.Num]; ok {
						if v, has := dObj["value"]; has {
							desc.Value = v
							desc.HasValue = true
						}
						if v, has := dObj["get"]; has {
							desc.Getter = v
							desc.HasGetter = true
						}
						if v, has := dObj["set"]; has {
							desc.Setter = v
							desc.HasSetter = true
						}
						desc.Enumerable = vm.jsToDescriptorBoolean(dObj["enumerable"], false)
						desc.Configurable = vm.jsToDescriptorBoolean(dObj["configurable"], false)
						desc.Writable = vm.jsToDescriptorBoolean(dObj["writable"], false)
					}
				}
				vm.jsSetDescriptor(objID, name, desc)
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
				if len(args) < 2 || (args[0].Type != VTJSObject && args[0].Type != VTJSFunction) {
					return Value{Type: VTJSUndefined}, true
				}
				name := vm.valueToString(args[1])
				desc, ok := vm.jsGetDescriptor(args[0].Num, name)
				if !ok {
					return Value{Type: VTJSUndefined}, true
				}
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
				return Value{Type: VTJSObject, Num: descObjID}, true
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
		case "RegExp":
			if strings.EqualFold(member, "test") {
				pattern := vm.jsObjectStringProperty(target, "pattern")
				flags := vm.jsObjectStringProperty(target, "flags")
				needle := ""
				if len(args) > 0 {
					needle = vm.valueToString(args[0])
				}
				rePattern := pattern
				if strings.Contains(strings.ToLower(flags), "i") {
					rePattern = "(?i)" + rePattern
				}
				re, err := regexp.Compile(rePattern)
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
		}
	}

	return Value{Type: VTJSUndefined}, false
}

// jsArgOrUndefined returns args[idx] or undefined for missing arguments.
func jsArgOrUndefined(args []Value, idx int) Value {
	if idx >= 0 && idx < len(args) {
		return args[idx]
	}
	return Value{Type: VTJSUndefined}
}

// jsStringReplace implements String.prototype.replace and replaceAll with size guards.
func (vm *VM) jsStringReplace(source string, patternArg Value, replacement string, replaceAll bool) Value {
	if patternArg.Type == VTJSObject {
		objType := vm.jsObjectStringProperty(patternArg, "__js_type")
		if objType == "RegExp" {
			pattern := vm.jsObjectStringProperty(patternArg, "pattern")
			flags := vm.jsObjectStringProperty(patternArg, "flags")
			return vm.jsStringReplaceRegex(source, pattern, flags, replacement, replaceAll)
		}
	}

	search := vm.valueToString(patternArg)
	if search == "" {
		if replaceAll {
			parts := len(source) + 1
			total := len(source) + parts*len(replacement)
			if !vm.jsEnsureStringSize(total) || !vm.jsChargeStringWork(total) {
				return Value{Type: VTJSUndefined}
			}
			var b strings.Builder
			b.Grow(total)
			b.WriteString(replacement)
			for i := 0; i < len(source); i++ {
				b.WriteByte(source[i])
				b.WriteString(replacement)
			}
			return NewString(b.String())
		}
		total := len(source) + len(replacement)
		if !vm.jsEnsureStringSize(total) || !vm.jsChargeStringWork(total) {
			return Value{Type: VTJSUndefined}
		}
		return NewString(replacement + source)
	}

	count := 1
	if replaceAll {
		count = -1
	}
	out := strings.Replace(source, search, replacement, count)
	if !vm.jsEnsureStringSize(len(out)) || !vm.jsChargeStringWork(len(out)) {
		return Value{Type: VTJSUndefined}
	}
	return NewString(out)
}

// jsStringReplaceRegex applies one or all RegExp matches to the source string.
func (vm *VM) jsStringReplaceRegex(source string, pattern string, flags string, replacement string, replaceAll bool) Value {
	rePattern := pattern
	flagsLower := strings.ToLower(flags)
	if strings.Contains(flagsLower, "i") {
		rePattern = "(?i)" + rePattern
	}
	re, err := regexp.Compile(rePattern)
	if err != nil {
		return NewString(source)
	}
	useAll := replaceAll || strings.Contains(flagsLower, "g")
	out := source
	if useAll {
		out = re.ReplaceAllString(source, replacement)
	} else {
		loc := re.FindStringIndex(source)
		if len(loc) == 2 {
			out = source[:loc[0]] + re.ReplaceAllString(source[loc[0]:loc[1]], replacement) + source[loc[1]:]
		}
	}
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
	switch target.Type {
	case VTNativeObject:
		vm.dispatchMemberSet(target.Num, member, val)
	case VTJSObject, VTJSFunction:
		var targetID int64
		if target.Type == VTJSObject || target.Type == VTJSFunction {
			targetID = target.Num
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
			return
		}
		obj[member] = val
		vm.jsSetDescriptor(targetID, member, jsDefaultPropertyDescriptor(val))
	}
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

func (vm *VM) jsCall(callee Value, thisVal Value, args []Value) Value {
	switch callee.Type {
	case VTJSFunction:
		if closure, ok := vm.jsFunctionItems[callee.Num]; ok && closure != nil && closure.isBound {
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
		if child.jsBeginFunctionCall(callee, thisVal, args, Value{Type: VTJSUndefined}, false) {
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
		case "Date":
			return NewString(time.Now().In(builtinCurrentLocation(vm)).Format("Mon Jan 02 2006 15:04:05 GMT-0700"))
		case "RegExp":
			if len(args) > 0 {
				return vm.jsNew(callee, args)
			}
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

// jsToNumber converts a Value to a numeric value (VTDouble) following JScript semantics.
func (vm *VM) jsToNumber(v Value) Value {
	switch v.Type {
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
	return vm.jsAddValues(a, b)
}

// jsSubtract implements JScript '-' operator (numeric subtraction).
func (vm *VM) jsSubtract(a Value, b Value) Value {
	aNum := vm.jsToNumber(a).Flt
	bNum := vm.jsToNumber(b).Flt
	return NewDouble(aNum - bNum)
}

// jsMultiply implements JScript '*' operator.
func (vm *VM) jsMultiply(a Value, b Value) Value {
	aNum := vm.jsToNumber(a).Flt
	bNum := vm.jsToNumber(b).Flt
	return NewDouble(aNum * bNum)
}

// jsDivide implements JScript '/' operator.
func (vm *VM) jsDivide(a Value, b Value) Value {
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
	aNum := vm.jsToNumber(a).Flt
	bNum := vm.jsToNumber(b).Flt
	if bNum == 0 || math.IsNaN(aNum) || math.IsNaN(bNum) || math.IsInf(aNum, 0) || math.IsInf(bNum, 0) {
		return NewDouble(math.NaN()) // NaN for modulo by zero
	}
	return NewDouble(math.Mod(aNum, bNum))
}

// jsNegate implements JScript unary '-' operator.
func (vm *VM) jsNegate(v Value) Value {
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
	aNum := vm.jsToNumber(a).Flt
	bNum := vm.jsToNumber(b).Flt
	return NewBool(aNum < bNum)
}

// jsGreater implements JScript '>' operator.
func (vm *VM) jsGreater(a Value, b Value) Value {
	aNum := vm.jsToNumber(a).Flt
	bNum := vm.jsToNumber(b).Flt
	return NewBool(aNum > bNum)
}

// jsLessEqual implements JScript '<=' operator.
func (vm *VM) jsLessEqual(a Value, b Value) Value {
	aNum := vm.jsToNumber(a).Flt
	bNum := vm.jsToNumber(b).Flt
	return NewBool(aNum <= bNum)
}

// jsGreaterEqual implements JScript '>=' operator.
func (vm *VM) jsGreaterEqual(a Value, b Value) Value {
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

// jsNew implements the 'new' operator for constructor calls.
func (vm *VM) jsNew(constructor Value, args []Value) Value {
	if constructor.Type == VTJSFunction {
		instanceID := vm.allocJSID()
		vm.jsObjectItems[instanceID] = make(map[string]Value, 8)
		vm.jsPropertyItems[instanceID] = make(map[string]jsPropertyDescriptor, 8)
		instance := Value{Type: VTJSObject, Num: instanceID}
		proto, deferred := vm.jsMemberGet(constructor, "prototype")
		if !deferred && proto.Type == VTJSObject {
			vm.jsObjectItems[instanceID]["__js_proto"] = proto
		} else {
			fallback := vm.jsGetIntrinsicPrototype("Object")
			if fallback.Type == VTJSObject {
				vm.jsObjectItems[instanceID]["__js_proto"] = fallback
			}
		}
		if vm.jsBeginFunctionCall(constructor, instance, args, instance, true) {
			return Value{Type: VTJSUndefined}
		}
		return instance
	}
	if constructor.Type == VTJSObject {
		ctorName := vm.jsObjectStringProperty(constructor, "__js_ctor")
		switch ctorName {
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
			return Value{Type: VTJSObject, Num: objID}
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
	switch arr.Type {
	case VTArray:
		if arr.Arr == nil {
			return Value{Type: VTJSUndefined}
		}
		indexNum := int(vm.jsToNumber(index).Flt)
		adjustedIndex := indexNum - arr.Arr.Lower
		if adjustedIndex < 0 || adjustedIndex >= len(arr.Arr.Values) {
			return Value{Type: VTJSUndefined}
		}
		return arr.Arr.Values[adjustedIndex]
	case VTJSObject:
		key := vm.valueToString(index)
		if v, _ := vm.jsMemberGet(arr, key); v.Type != VTJSUndefined {
			return v
		}
		return Value{Type: VTJSUndefined}
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
		key := vm.valueToString(index)
		vm.jsMemberSet(arr, key, value)
	}
}
