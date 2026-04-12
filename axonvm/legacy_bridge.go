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
	"strconv"
)

func legacyValueToInterface(v Value, vm *VM) interface{} {
	switch v.Type {
	case VTString:
		return v.String()
	case VTInteger:
		return int(v.Num) // PDF and XML expect int
	case VTDouble:
		return v.Flt
	case VTBool:
		return v.Num != 0
	case VTArray:
		if v.Arr != nil {
			var res []interface{}
			for _, item := range v.Arr.Values {
				res = append(res, legacyValueToInterface(item, vm))
			}
			return res
		}
	case VTNativeObject:
		if vm == nil {
			return nil
		}
		if obj, ok := vm.msxmlElementItems[v.Num]; ok {
			return obj
		}
		if obj, ok := vm.msxmlDOMItems[v.Num]; ok {
			return obj
		}
		if obj, ok := vm.msxmlNodeListItems[v.Num]; ok {
			return obj
		}
		if obj, ok := vm.msxmlParseErrorItems[v.Num]; ok {
			return obj
		}
		if obj, ok := vm.pdfItems[v.Num]; ok {
			return obj
		}
	}
	return nil
}

func legacyInterfaceToValue(i interface{}, vm *VM) Value {
	if i == nil {
		return NewEmpty()
	}
	switch v := i.(type) {
	case string:
		return NewString(v)
	case int:
		return NewInteger(int64(v))
	case int32:
		return NewInteger(int64(v))
	case int64:
		return NewInteger(v)
	case float64:
		return NewDouble(v)
	case float32:
		return NewDouble(float64(v))
	case bool:
		return NewBool(v)
	case []interface{}:
		var vals []Value
		for _, item := range v {
			vals = append(vals, legacyInterfaceToValue(item, vm))
		}
		return Value{Type: VTArray, Arr: NewVBArrayFromValues(0, vals)}
	case []byte:
		var vals []Value
		for _, b := range v {
			vals = append(vals, NewInteger(int64(b)))
		}
		return Value{Type: VTArray, Arr: NewVBArrayFromValues(0, vals)}
	case *MsXML2ServerXMLHTTP:
		id := vm.nextDynamicNativeID
		vm.nextDynamicNativeID++
		vm.msxmlServerItems[id] = v
		return Value{Type: VTNativeObject, Num: id}
	case *MsXML2DOMDocument:
		id := vm.nextDynamicNativeID
		vm.nextDynamicNativeID++
		vm.msxmlDOMItems[id] = v
		return Value{Type: VTNativeObject, Num: id}
	case *XMLNodeList:
		id := vm.nextDynamicNativeID
		vm.nextDynamicNativeID++
		vm.msxmlNodeListItems[id] = v
		return Value{Type: VTNativeObject, Num: id}
	case *ParseError:
		id := vm.nextDynamicNativeID
		vm.nextDynamicNativeID++
		vm.msxmlParseErrorItems[id] = v
		return Value{Type: VTNativeObject, Num: id}
	case *XMLElement:
		id := vm.nextDynamicNativeID
		vm.nextDynamicNativeID++
		vm.msxmlElementItems[id] = v
		return Value{Type: VTNativeObject, Num: id}
	case *G3PDF:
		id := vm.nextDynamicNativeID
		vm.nextDynamicNativeID++
		vm.pdfItems[id] = v
		return Value{Type: VTNativeObject, Num: id}
	case Value:
		return v
	}
	return NewString(fmt.Sprintf("%v", i))
}

func toInt(v interface{}) int {
	if v == nil {
		return 0
	}
	switch val := v.(type) {
	case int:
		return val
	case int64:
		return int(val)
	case float64:
		return int(val)
	case string:
		if i, err := strconv.Atoi(val); err == nil {
			return i
		}
	}
	return 0
}

func toFloat(v interface{}) float64 {
	if v == nil {
		return 0
	}
	switch val := v.(type) {
	case float64:
		return val
	case int:
		return float64(val)
	case int64:
		return float64(val)
	case string:
		if f, err := strconv.ParseFloat(val, 64); err == nil {
			return f
		}
	}
	return 0
}

func toBool(v interface{}) bool {
	if v == nil {
		return false
	}
	switch val := v.(type) {
	case bool:
		return val
	case int:
		return val != 0
	case float64:
		return val != 0
	case string:
		return val == "true" || val == "1" || val == "-1"
	}
	return false
}
