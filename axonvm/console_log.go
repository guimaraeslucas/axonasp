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
	"debug": {writer: os.Stdout, symbol: "⚙ ", level: "DEBUG", logFile: "console.log"},
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
	case "clear":
		// No-op for now in CLI
		return Value{Type: VTEmpty}
	}

	if len(args) == 0 {
		return Value{Type: VTEmpty}
	}

	format, supported := consoleMethodFormats[lower]
	if !supported {
		format = consoleMethodFormats["log"]
	}

	msg := consoleSerializeArgs(vm, args)
	if msg == "" && len(args) > 0 {
		msg = "(empty)"
	}
	consoleWriteMessage(vm, format, msg)

	return Value{Type: VTEmpty}
}

// jsCreateConsoleObject creates a JScript-compatible console object.
// This allows JScript code to see console.log, console.error, etc. as proper functions.
func (vm *VM) jsCreateConsoleObject() Value {
	objID := vm.allocJSID()
	obj := make(map[string]Value, 8)

	// Type marker
	obj["__js_type"] = NewString("console")

	// Helper to create a native function proxy
	createMethod := func(name string) Value {
		return vm.jsCreateIntrinsicFunction("console."+name, "Console"+strings.Title(name))
	}

	obj["log"] = createMethod("log")
	obj["warn"] = createMethod("warn")
	obj["error"] = createMethod("error")
	obj["info"] = createMethod("info")
	obj["debug"] = createMethod("debug")
	obj["trace"] = createMethod("trace")
	obj["clear"] = createMethod("clear")

	if proto := vm.jsGetIntrinsicPrototype("Object"); proto.Type == VTJSObject {
		obj["__js_proto"] = proto
	}

	vm.jsObjectItems[objID] = obj
	vm.jsPropertyItems[objID] = make(map[string]jsPropertyDescriptor, 8)

	return Value{Type: VTJSObject, Num: objID}
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
	for i := range args {
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
	if vm == nil || vm.host == nil {
		return format.writer, "\n"
	}

	// Always prefer host output in CLI mode if available, unless in TUI where we need more control.
	if vm.executionMode == ExecutionModeCLI || vm.executionMode == ExecutionModeTUI {
		if vm.host.Response() != nil && vm.host.Response().Output != nil {
			lineEnding := "\n"
			if vm.host.Request() != nil && vm.host.Request().ServerVars.Get("AXONASP_CLI_TUI") == "1" {
				lineEnding = ""
			}
			return vm.host.Response().Output, lineEnding
		}
	}

	return format.writer, "\n"
}

// consoleSerializeArg converts a single VM Value to a printable string.
// Primitive types are stringified directly. Arrays and objects are stringified recursively.
func consoleSerializeArg(vm *VM, v Value) string {
	return consoleSerializeValue(vm, v, make(map[int64]struct{}))
}

func consoleSerializeValue(vm *VM, v Value, visited map[int64]struct{}) string {
	if v.Type == VTArgRef && vm != nil {
		v = vm.stack[int(v.Num)]
	}

	switch v.Type {
	case VTArray:
		if v.Arr == nil {
			return "[]"
		}
		if _, ok := visited[v.Num]; ok {
			return "[Circular Array]"
		}
		visited[v.Num] = struct{}{}
		defer delete(visited, v.Num)

		items := make([]string, len(v.Arr.Values))
		for i, item := range v.Arr.Values {
			items[i] = consoleSerializeValue(vm, item, visited)
		}
		return "[" + strings.Join(items, ", ") + "]"

	case VTJSObject, VTJSFunction:
		if _, ok := visited[v.Num]; ok {
			return "[Circular Object]"
		}
		visited[v.Num] = struct{}{}
		defer delete(visited, v.Num)

		obj, ok := vm.jsObjectItems[v.Num]
		if !ok || obj == nil {
			return "{}"
		}

		jsType := vm.jsObjectStringProperty(v, "__js_type")
		if jsType == "" {
			jsType = vm.jsObjectStringProperty(v, "__js_ctor")
		}

		// Handle simple wrapper objects or well-known types as strings if they are small
		if jsType != "" && jsType != "Object" && jsType != "Array" && jsType != "Function" {
			// For console, process, etc., just show the type tag unless it's the root of log
			if len(visited) > 1 {
				return "[object " + jsType + "]"
			}
		}

		m := make([]string, 0, len(obj))
		for k, val := range obj {
			if strings.HasPrefix(k, "__js_") {
				continue // Skip internal markers
			}
			m = append(m, k+": "+consoleSerializeValue(vm, val, visited))
		}
		if jsType != "" && jsType != "Object" && jsType != "Function" {
			return jsType + " {" + strings.Join(m, ", ") + "}"
		}
		return "{" + strings.Join(m, ", ") + "}"

	case VTNativeObject:
		if vm != nil {
			if v.Num == nativeObjectConsole {
				return "[object console]"
			}
			if v.Num == nativeObjectResponse {
				return "[object Response]"
			}
			if v.Num == nativeObjectRequest {
				return "[object Request]"
			}
			if v.Num == nativeObjectServer {
				return "[object Server]"
			}
			if v.Num == nativeObjectSession {
				return "[object Session]"
			}
			if v.Num == nativeObjectApplication {
				return "[object Application]"
			}
			if _, ok := vm.dictionaryItems[v.Num]; ok {
				return consoleSerializeDictionary(vm, v, visited)
			}
		}
		return "[object NativeObject]"

	default:
		if vm != nil {
			return vm.valueToString(v)
		}
		return v.String()
	}
}

// consoleSerializeDictionary converts a native Dictionary object into a string.
func consoleSerializeDictionary(vm *VM, v Value, visited map[int64]struct{}) string {
	if vm == nil {
		return "{}"
	}
	keysVal, _ := vm.dispatchDictionaryMethod(v.Num, "Keys", nil)
	itemsVal, _ := vm.dispatchDictionaryMethod(v.Num, "Items", nil)
	if keysVal.Type != VTArray || itemsVal.Type != VTArray ||
		keysVal.Arr == nil || itemsVal.Arr == nil {
		return "{}"
	}
	m := make([]string, len(keysVal.Arr.Values))
	for i := 0; i < len(keysVal.Arr.Values) && i < len(itemsVal.Arr.Values); i++ {
		k := keysVal.Arr.Values[i].String()
		m[i] = k + ": " + consoleSerializeValue(vm, itemsVal.Arr.Values[i], visited)
	}
	return "Dictionary {" + strings.Join(m, ", ") + "}"
}

// consoleValueToInterface recursively converts a VM Value to a Go interface{} for JSON marshaling.
func consoleValueToInterface(vm *VM, v Value) any {
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
			return []any{}
		}
		items := make([]any, len(v.Arr.Values))
		for i, item := range v.Arr.Values {
			items[i] = consoleValueToInterface(vm, item)
		}
		return items
	case VTJSObject:
		if vm != nil {
			if obj, ok := vm.jsObjectItems[v.Num]; ok && obj != nil {
				result := make(map[string]any, len(obj))
				for k, val := range obj {
					if strings.HasPrefix(k, "__js_") {
						continue
					}
					result[k] = consoleValueToInterface(vm, val)
				}
				return result
			}
		}
		return map[string]any{}
	case VTNativeObject:
		if vm != nil {
			switch v.Num {
			case nativeObjectConsole:
				return "[object console]"
			case nativeObjectResponse:
				return "[object Response]"
			case nativeObjectRequest:
				return "[object Request]"
			case nativeObjectServer:
				return "[object Server]"
			case nativeObjectSession:
				return "[object Session]"
			case nativeObjectApplication:
				return "[object Application]"
			}
			if _, ok := vm.dictionaryItems[v.Num]; ok {
				keysVal, _ := vm.dispatchDictionaryMethod(v.Num, "Keys", nil)
				itemsVal, _ := vm.dispatchDictionaryMethod(v.Num, "Items", nil)
				if keysVal.Type == VTArray && itemsVal.Type == VTArray &&
					keysVal.Arr != nil && itemsVal.Arr != nil {
					m := make(map[string]any, len(keysVal.Arr.Values))
					for i := 0; i < len(keysVal.Arr.Values) && i < len(itemsVal.Arr.Values); i++ {
						k := keysVal.Arr.Values[i].String()
						m[k] = consoleValueToInterface(vm, itemsVal.Arr.Values[i])
					}
					return m
				}
			}
		}
		return "[object NativeObject]"
	default:
		if vm != nil {
			return vm.valueToString(v)
		}
		return v.String()
	}
}
