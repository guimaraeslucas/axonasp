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

type jsEnvFrame struct {
	parentID int64
	bindings map[string]Value
}

type jsFunctionObject struct {
	name    string
	params  []string
	startIP int
	endIP   int
	envID   int64
}

type jsCallFrame struct {
	returnIP int
	envID    int64
	thisVal  Value
	tryDepth int
	savedSP  int
}

type jsForInEnumerator struct {
	keys  []string
	index int
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
	bindings["NaN"] = NewDouble(math.NaN())
	bindings["Infinity"] = NewDouble(math.Inf(1))
	bindings["undefined"] = Value{Type: VTJSUndefined}
	bindings["isNaN"] = vm.jsCreateIntrinsicObject("", "isNaN")
	bindings["isFinite"] = vm.jsCreateIntrinsicObject("", "isFinite")
	bindings["parseInt"] = vm.jsCreateIntrinsicObject("", "parseInt")
	bindings["parseFloat"] = vm.jsCreateIntrinsicObject("", "parseFloat")
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
	// (globalNames/globalNameIndex). In that case, expose builtins by their
	// canonical fixed slots to keep JScript global lookup consistent.
	if builtinIdx, ok := BuiltinIndex[lowerName]; ok {
		globalIdx := 9 + builtinIdx // 7 intrinsics + 2 transaction handlers
		if globalIdx >= 0 && globalIdx < len(vm.Globals) && vm.Globals[globalIdx].Type == VTBuiltin {
			return globalIdx, true
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
		sa := vm.valueToString(a)
		sb := vm.valueToString(b)
		total := len(sa) + len(sb)
		if !vm.jsEnsureStringSize(total) || !vm.jsChargeStringWork(total) {
			return Value{Type: VTJSUndefined}
		}
		return NewString(sa + sb)
	}
	return NewDouble(vm.jsToNumber(a).Flt + vm.jsToNumber(b).Flt)
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
	vm.jsFunctionItems[id] = &jsFunctionObject{
		name:    template.Str,
		params:  append([]string(nil), template.Names...),
		startIP: int(template.Num),
		endIP:   int(template.Flt),
		envID:   vm.jsActiveEnvID,
	}
	return Value{Type: VTJSFunction, Num: id}
}

func (vm *VM) jsBeginFunctionCall(fn Value, thisVal Value, args []Value) bool {
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
	vm.push(retVal)
}

func (vm *VM) jsMemberGet(target Value, member string) Value {
	switch target.Type {
	case VTNativeObject:
		return vm.dispatchMemberGet(target, member)
	case VTString:
		if strings.EqualFold(member, "length") {
			return NewInteger(int64(len(target.Str)))
		}
		return Value{Type: VTJSUndefined}
	case VTArray:
		if target.Arr != nil && strings.EqualFold(member, "length") {
			return NewInteger(int64(len(target.Arr.Values)))
		}
		return Value{Type: VTJSUndefined}
	case VTDate:
		return Value{Type: VTJSUndefined}
	case VTJSObject:
		if obj, ok := vm.jsObjectItems[target.Num]; ok {
			if val, exists := obj[member]; exists {
				return val
			}
		}
		return Value{Type: VTJSUndefined}
	default:
		return Value{Type: VTJSUndefined}
	}
}

func (vm *VM) jsCallMember(target Value, member string, args []Value) (Value, bool) {
	switch target.Type {
	case VTString:
		text := target.Str
		switch {
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
		}
	case VTArray:
		if target.Arr == nil {
			return Value{Type: VTJSUndefined}, true
		}
		switch {
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
		case strings.EqualFold(member, "valueOf"):
			t := valueToTimeInLocale(vm, target)
			return NewInteger(t.UnixNano() / int64(time.Millisecond)), true
		}
	case VTJSObject:
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
		countVal := vm.dispatchMemberGet(source, "Count")
		return int(vm.jsToNumber(countVal).Flt)
	default:
		return 0
	}
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
		zeroBased := vm.dispatchNativeCall(source.Num, "Item", []Value{NewInteger(int64(idx))})
		if zeroBased.Type != VTJSUndefined && zeroBased.Type != VTEmpty {
			return zeroBased
		}
		return vm.dispatchNativeCall(source.Num, "Item", []Value{NewInteger(int64(idx + 1))})
	default:
		return Value{Type: VTJSUndefined}
	}
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
	dim := 1
	cur := arr
	for cur != nil && len(cur.Values) > 0 {
		next, nextOK := toVBArray(cur.Values[0])
		if !nextOK {
			break
		}
		dim++
		cur = next
	}
	return dim
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
	case VTJSObject:
		obj, ok := vm.jsObjectItems[target.Num]
		if !ok {
			obj = make(map[string]Value, 8)
			vm.jsObjectItems[target.Num] = obj
		}
		obj[member] = val
	}
}

func (vm *VM) jsEnumerateForInKeys(source Value) []string {
	if source.Type == VTJSObject {
		obj, ok := vm.jsObjectItems[source.Num]
		if !ok || len(obj) == 0 {
			return nil
		}
		keys := make([]string, 0, len(obj))
		for k := range obj {
			keys = append(keys, k)
		}
		sort.Strings(keys)
		return keys
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
		if vm.jsBeginFunctionCall(callee, thisVal, args) {
			return Value{Type: VTJSUndefined}
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
			if len(args) == 0 {
				return NewDouble(math.NaN())
			}
			s := strings.TrimSpace(vm.valueToString(args[0]))
			if s == "" {
				return NewDouble(math.NaN())
			}
			radix := 10
			if len(args) > 1 {
				radix = int(vm.jsToNumber(args[1]).Flt)
			}
			// Simplified parseInt - for production should be more robust
			if radix == 0 {
				radix = 10
			}
			val, err := strconv.ParseInt(s, radix, 64)
			if err != nil {
				// try floating point then truncate
				f, err2 := strconv.ParseFloat(s, 64)
				if err2 == nil {
					return NewDouble(math.Trunc(f))
				}
				return NewDouble(math.NaN())
			}
			return NewDouble(float64(val))
		case "parseFloat":
			if len(args) == 0 {
				return NewDouble(math.NaN())
			}
			s := strings.TrimSpace(vm.valueToString(args[0]))
			val, err := strconv.ParseFloat(s, 64)
			if err != nil {
				return NewDouble(math.NaN())
			}
			return NewDouble(val)
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
	// TODO: Implement full member index access
	// For now, return undefined
	return Value{Type: VTJSUndefined}
}

// jsMemberIndexSet implements member[index] = value assignment.
func (vm *VM) jsMemberIndexSet(obj Value, index Value, value Value, memberName string) {
	// TODO: Implement full member index assignment
}

// jsNew implements the 'new' operator for constructor calls.
func (vm *VM) jsNew(constructor Value, args []Value) Value {
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
			return Value{Type: VTJSObject, Num: objID}
		}
	}

	objID := vm.allocJSID()
	vm.jsObjectItems[objID] = make(map[string]Value, 8)
	return Value{Type: VTJSObject, Num: objID}
}

// jsMemberDelete implements the 'delete' operator for object properties.
func (vm *VM) jsMemberDelete(obj Value, member string) bool {
	if obj.Type == VTJSObject {
		if jsObj, ok := vm.jsObjectItems[obj.Num]; ok {
			delete(jsObj, member)
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
		obj, ok := vm.jsObjectItems[arr.Num]
		if !ok {
			return Value{Type: VTJSUndefined}
		}
		key := vm.valueToString(index)
		if v, exists := obj[key]; exists {
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
		obj, ok := vm.jsObjectItems[arr.Num]
		if !ok {
			obj = make(map[string]Value, 8)
			vm.jsObjectItems[arr.Num] = obj
		}
		key := vm.valueToString(index)
		obj[key] = value
	}
}
