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
	"io"
	"os"
	"strings"
	"time"
)

// consoleOutputFormat defines a specific console method's output destination and display symbol.
type consoleOutputFormat struct {
	writer  io.Writer
	symbol  string
	level   string
	logFile string // "console.log" or "error.log"
}

// consoleMethodFormats maps lowercased method names to their output routing configuration.
// The decorative unicode symbol is written to the stream only; the log file receives plain text.
var consoleMethodFormats = map[string]consoleOutputFormat{
	"log":   {writer: os.Stdout, symbol: "⌨ ", level: "LOG", logFile: "console.log"},
	"info":  {writer: os.Stdout, symbol: "ℹ ", level: "INFO", logFile: "console.log"},
	"error": {writer: os.Stderr, symbol: "✖ ", level: "ERROR", logFile: "error.log"},
	"warn":  {writer: os.Stderr, symbol: "⚠ ", level: "WARN", logFile: "error.log"},
}

// consoleDispatch is the entry point for all console.method(args) calls from both
// VBScript (via dispatchNativeCall) and JScript (via OpJSCallMember → dispatchNativeCall).
// It formats the first argument into a printable string, writes to the correct stream,
// and conditionally appends a clean (no-symbol) entry to the appropriate log file.
func consoleDispatch(vm *VM, method string, args []Value) Value {
	if len(args) == 0 {
		return Value{Type: VTEmpty}
	}

	lower := strings.ToLower(method)
	format, supported := consoleMethodFormats[lower]
	if !supported {
		return Value{Type: VTEmpty}
	}

	msg := consoleSerializeArg(vm, args[0])
	timestamp := time.Now().Format("2006/01/02 15:04:05")

	// Write decorated output to the target stream (stdout or stderr).
	fmt.Fprintf(format.writer, "%s [%s] %s %s\n", timestamp, format.level, format.symbol, msg)

	// Write a plain (symbol-free) entry to the configured log file.
	writeConsoleLogToFile(format.logFile, format.level, msg, timestamp)

	return Value{Type: VTEmpty}
}

// consoleSerializeArg converts a single VM Value to a printable string.
// Primitive types are stringified directly. Arrays and objects are JSON-encoded.
func consoleSerializeArg(vm *VM, v Value) string {
	switch v.Type {
	case VTArray:
		return consoleSerializeArray(vm, v.Arr)
	case VTJSObject:
		return consoleSerializeJSObject(vm, v)
	case VTNativeObject:
		// Attempt to serialize a Scripting.Dictionary as a JSON object.
		if vm != nil {
			if _, ok := vm.dictionaryItems[v.Num]; ok {
				return consoleSerializeDictionary(vm, v)
			}
		}
		return "[object]"
	default:
		if vm != nil {
			return vm.valueToString(v)
		}
		return v.String()
	}
}

// consoleSerializeArray converts a VBScript or JScript array into a JSON array string.
func consoleSerializeArray(vm *VM, arr *VBArray) string {
	if arr == nil || len(arr.Values) == 0 {
		return "[]"
	}
	items := make([]interface{}, len(arr.Values))
	for i, item := range arr.Values {
		items[i] = consoleValueToInterface(vm, item)
	}
	b, err := json.Marshal(items)
	if err != nil {
		return "[]"
	}
	return string(b)
}

// consoleSerializeJSObject converts a JScript object (VTJSObject) into a JSON object string.
func consoleSerializeJSObject(vm *VM, v Value) string {
	if vm == nil {
		return "{}"
	}
	obj, ok := vm.jsObjectItems[v.Num]
	if !ok || obj == nil {
		return "{}"
	}
	m := make(map[string]interface{}, len(obj))
	for k, val := range obj {
		m[k] = consoleValueToInterface(vm, val)
	}
	b, err := json.Marshal(m)
	if err != nil {
		return "{}"
	}
	return string(b)
}

// consoleSerializeDictionary converts a native Dictionary object into a JSON object string.
func consoleSerializeDictionary(vm *VM, v Value) string {
	if vm == nil {
		return "{}"
	}
	keysVal, _ := vm.dispatchDictionaryMethod(v.Num, "Keys", nil)
	itemsVal, _ := vm.dispatchDictionaryMethod(v.Num, "Items", nil)
	if keysVal.Type != VTArray || itemsVal.Type != VTArray ||
		keysVal.Arr == nil || itemsVal.Arr == nil {
		return "{}"
	}
	m := make(map[string]interface{}, len(keysVal.Arr.Values))
	for i := 0; i < len(keysVal.Arr.Values) && i < len(itemsVal.Arr.Values); i++ {
		k := keysVal.Arr.Values[i].String()
		m[k] = consoleValueToInterface(vm, itemsVal.Arr.Values[i])
	}
	b, err := json.Marshal(m)
	if err != nil {
		return "{}"
	}
	return string(b)
}

// consoleValueToInterface recursively converts a VM Value to a Go interface{} for JSON marshaling.
func consoleValueToInterface(vm *VM, v Value) interface{} {
	switch v.Type {
	case VTBool:
		return v.Num != 0
	case VTInteger:
		return v.Num
	case VTDouble:
		return v.Flt
	case VTString:
		return v.Str
	case VTNull:
		return nil
	case VTEmpty:
		return nil
	case VTArray:
		if v.Arr == nil {
			return []interface{}{}
		}
		items := make([]interface{}, len(v.Arr.Values))
		for i, item := range v.Arr.Values {
			items[i] = consoleValueToInterface(vm, item)
		}
		return items
	case VTJSObject:
		if vm != nil {
			if obj, ok := vm.jsObjectItems[v.Num]; ok && obj != nil {
				m := make(map[string]interface{}, len(obj))
				for k, val := range obj {
					m[k] = consoleValueToInterface(vm, val)
				}
				return m
			}
		}
		return map[string]interface{}{}
	default:
		if vm != nil {
			return vm.valueToString(v)
		}
		return v.String()
	}
}
