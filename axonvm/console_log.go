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
	"strconv"
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
	"err":   {writer: os.Stderr, symbol: "✖ ", level: "ERROR", logFile: "error.log"},
	"error": {writer: os.Stderr, symbol: "✖ ", level: "ERROR", logFile: "error.log"},
	"warn":  {writer: os.Stderr, symbol: "⚠ ", level: "WARN", logFile: "error.log"},
	"trace": {writer: os.Stderr, symbol: "↳ ", level: "TRACE", logFile: "error.log"},
}

const consoleDefaultTimerLabel = "default"

// consoleDispatch is the entry point for all console.method(args) calls from both
// VBScript (via dispatchNativeCall) and JScript (via OpJSCallMember → dispatchNativeCall).
// It formats the first argument into a printable string, writes to the correct stream,
// and conditionally appends a clean (no-symbol) entry to the appropriate log file.
func consoleDispatch(vm *VM, method string, args []Value) Value {
	lower := strings.ToLower(method)
	switch lower {
	case "time":
		label := consoleTimerLabel(vm, args)
		if vm != nil {
			if vm.consoleTimerItems == nil {
				vm.consoleTimerItems = make(map[string]time.Time)
			}
			vm.consoleTimerItems[label] = time.Now()
		}
		return Value{Type: VTEmpty}
	case "timeend":
		label := consoleTimerLabel(vm, args)
		if vm == nil || vm.consoleTimerItems == nil {
			consoleWriteMessage(vm, consoleMethodFormats["warn"], "No such label '"+label+"' for console.timeEnd()")
			return Value{Type: VTEmpty}
		}
		startedAt, exists := vm.consoleTimerItems[label]
		if !exists {
			consoleWriteMessage(vm, consoleMethodFormats["warn"], "No such label '"+label+"' for console.timeEnd()")
			return Value{Type: VTEmpty}
		}
		delete(vm.consoleTimerItems, label)
		elapsedMs := float64(time.Since(startedAt).Nanoseconds()) / 1e6
		consoleWriteMessage(vm, consoleMethodFormats["log"], fmt.Sprintf("%s: %.3fms", label, elapsedMs))
		return Value{Type: VTEmpty}
	case "dir":
		msg := "undefined"
		if len(args) > 0 {
			msg = consoleInspectArg(vm, args[0])
		}
		consoleWriteMessage(vm, consoleMethodFormats["log"], msg)
		return Value{Type: VTEmpty}
	case "trace":
		if !consoleIsJScriptContext(vm) {
			return Value{Type: VTEmpty}
		}
		message := "Trace"
		if len(args) > 0 {
			message = "Trace: " + consoleSerializeArg(vm, args[0])
		}
		stack := consoleBuildJSTrace(vm)
		if stack != "" {
			message += "\n" + stack
		}
		consoleWriteMessage(vm, consoleMethodFormats["trace"], message)
		return Value{Type: VTEmpty}
	}

	if len(args) == 0 {
		return Value{Type: VTEmpty}
	}

	format, supported := consoleMethodFormats[lower]
	if !supported {
		return Value{Type: VTEmpty}
	}

	msg := consoleSerializeArgs(vm, args)
	consoleWriteMessage(vm, format, msg)

	return Value{Type: VTEmpty}
}

// consoleSerializeArgs converts all console arguments into one printable message.
// Arguments are joined by a single space to match JavaScript console behavior.
func consoleSerializeArgs(vm *VM, args []Value) string {
	if len(args) == 0 {
		return ""
	}
	if len(args) == 1 {
		return consoleSerializeArg(vm, args[0])
	}

	parts := make([]string, len(args))
	for i := 0; i < len(args); i++ {
		parts[i] = consoleSerializeArg(vm, args[i])
	}
	return strings.Join(parts, " ")
}

// consoleWriteMessage writes one decorated line to console output and a plain line to log files.
func consoleWriteMessage(vm *VM, format consoleOutputFormat, msg string) {
	timestamp := time.Now().Format("2006/01/02 15:04:05")
	writer, lineEnding := resolveConsoleOutputTarget(vm, format)

	// Write decorated output to the target stream (stdout or stderr).
	fmt.Fprintf(writer, "%s [%s] %s %s%s", timestamp, format.level, format.symbol, msg, lineEnding)

	// Write a plain (symbol-free) entry to the configured log file.
	writeConsoleLogToFile(format.logFile, format.level, msg, timestamp)
}

// consoleTimerLabel returns the normalized timer label, defaulting to "default".
func consoleTimerLabel(vm *VM, args []Value) string {
	if len(args) == 0 {
		return consoleDefaultTimerLabel
	}
	label := strings.TrimSpace(consoleSerializeArg(vm, args[0]))
	if label == "" {
		return consoleDefaultTimerLabel
	}
	return label
}

// consoleIsJScriptContext reports whether the current execution is in the JScript runtime.
func consoleIsJScriptContext(vm *VM) bool {
	if vm == nil {
		return false
	}
	if vm.engineMode == EngineModeJavaScript {
		return true
	}
	return len(vm.jsCallStack) > 0 || vm.jsActiveEnvID != 0 || vm.jsRootEnvID != 0
}

// consoleInspectArg renders one value in a multiline, inspection-focused format.
func consoleInspectArg(vm *VM, v Value) string {
	inspected := consoleValueToInterface(vm, v)
	b, err := json.MarshalIndent(inspected, "", "  ")
	if err != nil {
		return consoleSerializeArg(vm, v)
	}
	return string(b)
}

// consoleBuildJSTrace formats the active JScript call stack using file/line/column metadata.
func consoleBuildJSTrace(vm *VM) string {
	if vm == nil {
		return ""
	}
	var b strings.Builder

	file := vm.sourceName
	if strings.TrimSpace(file) == "" {
		file = vm.baseSourceName
	}
	if strings.TrimSpace(file) == "" {
		file = "<script>"
	}

	line := vm.lastLine
	if line <= 0 {
		line = 1
	}
	col := vm.lastColumn
	if col <= 0 {
		col = 1
	}

	frameName := "<global>"
	if n := len(vm.jsCallStack); n > 0 {
		frameName = consoleJSFrameName(vm, vm.jsCallStack[n-1])
	}
	b.WriteString("    at ")
	b.WriteString(frameName)
	b.WriteString(" (")
	b.WriteString(file)
	b.WriteString(":")
	b.WriteString(strconv.Itoa(line))
	b.WriteString(":")
	b.WriteString(strconv.Itoa(col))
	b.WriteString(")")

	for i := len(vm.jsCallStack) - 1; i >= 0; i-- {
		if i == len(vm.jsCallStack)-1 {
			continue
		}
		frame := vm.jsCallStack[i]
		frameFile := frame.callFile
		if strings.TrimSpace(frameFile) == "" {
			frameFile = file
		}
		frameLine := frame.callLine
		if frameLine <= 0 {
			frameLine = 1
		}
		frameCol := frame.callColumn
		if frameCol <= 0 {
			frameCol = 1
		}
		b.WriteString("\n    at ")
		b.WriteString(consoleJSFrameName(vm, frame))
		b.WriteString(" (")
		b.WriteString(frameFile)
		b.WriteString(":")
		b.WriteString(strconv.Itoa(frameLine))
		b.WriteString(":")
		b.WriteString(strconv.Itoa(frameCol))
		b.WriteString(")")
	}

	return b.String()
}

// consoleJSFrameName resolves one JScript call frame function name for trace output.
func consoleJSFrameName(vm *VM, frame jsCallFrame) string {
	if vm == nil {
		return "<anonymous>"
	}
	if frame.fn.Type != VTJSFunction {
		return "<anonymous>"
	}
	closure, ok := vm.jsFunctionItems[frame.fn.Num]
	if !ok || closure == nil {
		return "<anonymous>"
	}
	if strings.TrimSpace(closure.name) == "" {
		return "<anonymous>"
	}
	return closure.name
}

// resolveConsoleOutputTarget selects where console output should be written.
// In CLI TUI mode we write to the host output buffer and avoid a trailing newline
// so the TUI output box remains stable. Other runtimes keep the default stream+newline behavior.
func resolveConsoleOutputTarget(vm *VM, format consoleOutputFormat) (io.Writer, string) {
	if vm == nil || vm.host == nil || vm.host.Request() == nil {
		return format.writer, "\n"
	}

	if vm.host.Request().ServerVars.Get("AXONASP_CLI_TUI") != "1" {
		return format.writer, "\n"
	}

	if vm.host.Response() != nil && vm.host.Response().Output != nil {
		return vm.host.Response().Output, ""
	}

	return vm.host, ""
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
