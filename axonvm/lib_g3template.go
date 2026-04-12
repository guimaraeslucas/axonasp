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
	"bytes"
	"html/template"
	"strings"
)

type G3Template struct {
	vm *VM
}

// newG3TemplateObject instantiates the G3Template custom functions library.
func (vm *VM) newG3TemplateObject() Value {
	obj := &G3Template{vm: vm}
	id := vm.nextDynamicNativeID
	vm.nextDynamicNativeID++
	vm.g3templateItems[id] = obj
	return Value{Type: VTNativeObject, Num: id}
}

// DispatchPropertyGet acts as a getter.
func (t *G3Template) DispatchPropertyGet(propertyName string) Value {
	return t.DispatchMethod(propertyName, nil)
}

// DispatchMethod provides O(1) string matching resolution for all custom template functions.
func (t *G3Template) DispatchMethod(methodName string, args []Value) Value {
	funcLower := strings.ToLower(methodName)

	switch funcLower {
	case "render":
		if len(args) < 1 {
			return NewString("Error: Template path required")
		}

		relPath := args[0].String()
		fullPath := relPath
		if t.vm.host != nil && t.vm.host.Server() != nil {
			fullPath = t.vm.host.Server().MapPath(relPath)
		}

		var data interface{}
		if len(args) > 1 {
			data = t.vmValueToGoValue(args[1])
		}

		tmpl, err := template.ParseFiles(fullPath)
		if err != nil {
			return NewString("Error parsing template: " + err.Error())
		}

		var buf bytes.Buffer
		if err := tmpl.Execute(&buf, data); err != nil {
			return NewString("Error executing template: " + err.Error())
		}

		return NewString(buf.String())
	}

	return NewEmpty()
}

// vmValueToGoValue converts VM types to Go types for html/template consumption
func (t *G3Template) vmValueToGoValue(v Value) interface{} {
	switch v.Type {
	case VTArray:
		if v.Arr == nil {
			return []interface{}{}
		}
		arr := make([]interface{}, len(v.Arr.Values))
		for i, item := range v.Arr.Values {
			arr[i] = t.vmValueToGoValue(item)
		}
		return arr
	case VTNativeObject:
		if _, ok := t.vm.dictionaryItems[v.Num]; ok {
			m := make(map[string]interface{})
			keysVal, _ := t.vm.dispatchDictionaryMethod(v.Num, "Keys", nil)
			itemsVal, _ := t.vm.dispatchDictionaryMethod(v.Num, "Items", nil)
			if keysVal.Type == VTArray && itemsVal.Type == VTArray && keysVal.Arr != nil && itemsVal.Arr != nil {
				for i := 0; i < len(keysVal.Arr.Values); i++ {
					k := keysVal.Arr.Values[i].String()
					m[k] = t.vmValueToGoValue(itemsVal.Arr.Values[i])
				}
			}
			return m
		}
		return nil
	case VTString:
		return v.String()
	case VTInteger:
		return v.Num
	case VTDouble:
		return v.Flt
	case VTBool:
		return v.Num != 0
	case VTNull, VTEmpty:
		return nil
	default:
		return v.String()
	}
}
