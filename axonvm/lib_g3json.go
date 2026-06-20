//go:build !lib_g3json_disabled

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
	"math"
	"os"
	"strings"
)

type G3JSON struct {
	vm *VM
}

// newG3JSONObject instantiates the G3JSON custom functions library.
func (vm *VM) newG3JSONObject() Value {
	obj := &G3JSON{vm: vm}
	id := vm.nextDynamicNativeID
	vm.nextDynamicNativeID++
	vm.g3jsonItems[id] = obj
	return Value{Type: VTNativeObject, Num: id}
}

// DispatchPropertyGet acts as a getter.
func (j *G3JSON) DispatchPropertyGet(propertyName string) Value {
	return j.DispatchMethod(propertyName, nil)
}

// DispatchMethod provides O(1) string matching resolution for all custom JSON functions.
func (j *G3JSON) DispatchMethod(methodName string, args []Value) Value {
	funcLower := strings.ToLower(methodName)

	switch funcLower {

	case "parse":
		if len(args) == 0 {
			return NewEmpty()
		}
		jsonStr := args[0].String()
		if jsonStr == "" {
			return NewEmpty()
		}
		var result any
		err := json.Unmarshal([]byte(jsonStr), &result)
		if err != nil {
			return NewEmpty()
		}
		return j.goValueToVMValue(result)

	case "stringify":
		if len(args) == 0 {
			return NewString("")
		}
		goVal := j.vmValueToGoValue(args[0], make(map[int64]bool), make(map[*VBArray]bool))
		bytes, err := json.Marshal(goVal)
		if err != nil {
			return NewString("")
		}
		return NewString(string(bytes))

	case "newobject":
		return j.vm.newDictionaryObject()

	case "newarray":
		return Value{Type: VTArray, Arr: NewVBArrayFromValues(0, []Value{})}

	case "loadfile":
		if len(args) == 0 {
			return NewEmpty()
		}
		path := args[0].String()
		if j.vm.host != nil && j.vm.host.Server() != nil {
			path = j.vm.host.Server().MapPath(path)
		}
		content, err := os.ReadFile(path)
		if err != nil {
			return NewEmpty()
		}
		var result any
		err = json.Unmarshal(content, &result)
		if err != nil {
			return NewEmpty()
		}
		return j.goValueToVMValue(result)

	}

	return NewEmpty()
}

// goValueToVMValue recursively converts Go nested maps/slices into VBScript Dictionaries/Arrays
func (j *G3JSON) goValueToVMValue(data any) Value {
	if data == nil {
		return Value{Type: VTNull}
	}

	switch v := data.(type) {
	case map[string]any:
		dictVal := j.vm.newDictionaryObject()
		if _, ok := j.vm.dictionaryItems[dictVal.Num]; ok {
			for key, val := range v {
				// We use dispatchDictionaryMethod "Add" to populate safely
				j.vm.dispatchDictionaryMethod(dictVal.Num, "Add", []Value{NewString(key), j.goValueToVMValue(val)})
			}
		}
		return dictVal

	case []any:
		arr := make([]Value, len(v))
		for i, item := range v {
			arr[i] = j.goValueToVMValue(item)
		}
		return Value{Type: VTArray, Arr: NewVBArrayFromValues(0, arr)}

	case string:
		return NewString(v)
	case float64:
		// Check if it's an integer
		if v == float64(int64(v)) {
			return NewInteger(int64(v))
		}
		return NewDouble(v)
	case bool:
		return NewBool(v)
	default:
		return NewString(j.vm.valueToString(Value{})) // Empty fallback
	}
}

// vmValueToGoValue recursively converts VM values into json.Marshal-compatible
// Go values. It guards against native-object/array cycles and normalizes
// unsupported VM-only value kinds into JSON-safe fallbacks.
func (j *G3JSON) vmValueToGoValue(v Value, seenNative map[int64]bool, seenArrays map[*VBArray]bool) any {
	switch v.Type {
	case VTArray:
		if v.Arr == nil {
			return []any{}
		}
		if seenArrays[v.Arr] {
			return nil
		}
		seenArrays[v.Arr] = true
		defer delete(seenArrays, v.Arr)
		arr := make([]any, len(v.Arr.Values))
		for i, item := range v.Arr.Values {
			arr[i] = j.vmValueToGoValue(item, seenNative, seenArrays)
		}
		return arr

	case VTNativeObject:
		if seenNative[v.Num] {
			return nil
		}
		seenNative[v.Num] = true
		defer delete(seenNative, v.Num)

		if _, ok := j.vm.dictionaryItems[v.Num]; ok {
			m := make(map[string]any)
			// We can iterate the dictionary keys
			keysVal, _ := j.vm.dispatchDictionaryMethod(v.Num, "Keys", nil)
			itemsVal, _ := j.vm.dispatchDictionaryMethod(v.Num, "Items", nil)
			if keysVal.Type == VTArray && itemsVal.Type == VTArray && keysVal.Arr != nil && itemsVal.Arr != nil {
				limit := min(len(itemsVal.Arr.Values), len(keysVal.Arr.Values))
				for i := 0; i < limit; i++ {
					k := keysVal.Arr.Values[i].String()
					m[k] = j.vmValueToGoValue(itemsVal.Arr.Values[i], seenNative, seenArrays)
				}
			}
			return m
		}
		// Unknown native objects are represented as descriptive strings rather
		// than forcing null, preserving useful diagnostics in JSON outputs.
		return v.String()

	case VTObject:
		if v.Num == 0 {
			return nil
		}
		return v.String()

	case VTString:
		return v.String()
	case VTInteger:
		return v.Num
	case VTDouble:
		if math.IsNaN(v.Flt) || math.IsInf(v.Flt, 0) {
			return nil
		}
		return v.Flt
	case VTBool:
		return v.Num != 0
	case VTDate:
		return v.String()
	case VTNothing:
		return nil
	case VTNull, VTEmpty:
		return nil
	case VTBuiltin, VTUserSub, VTArgRef:
		return v.String()
	case VTJSUndefined:
		return nil
	case VTJSObject, VTJSFunction, VTJSFunctionTemplate, VTJSArrowFunctionTemplate, VTSymbol:
		return v.String()
	default:
		return v.String()
	}
}
