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
	"strings"
	"testing"
	"time"
)

// TestResolveConsoleOutputTarget_CLITUI verifies that TUI mode routes console output
// to the host output buffer and suppresses trailing newlines.
func TestResolveConsoleOutputTarget_CLITUI(t *testing.T) {
	host := NewMockHost()
	var out bytes.Buffer
	host.SetOutput(&out)
	host.Request().ServerVars.Add("AXONASP_CLI_TUI", "1")

	vm := NewVM(nil, nil, 0)
	vm.SetHost(host)

	writer, lineEnding := resolveConsoleOutputTarget(vm, consoleMethodFormats["log"])
	if lineEnding != "" {
		t.Fatalf("expected no line ending in TUI mode, got %q", lineEnding)
	}
	if writer != &out {
		t.Fatalf("expected TUI writer to be host output buffer")
	}

	consoleDispatch(vm, "log", []Value{NewString("hello")})
	rendered := out.String()
	if rendered == "" {
		t.Fatalf("expected console output in host buffer")
	}
	if strings.HasSuffix(rendered, "\n") {
		t.Fatalf("expected no trailing newline in TUI mode output, got %q", rendered)
	}
}

// TestResolveConsoleOutputTarget_Default verifies non-TUI behavior keeps newline output.
func TestResolveConsoleOutputTarget_Default(t *testing.T) {
	host := NewMockHost()
	vm := NewVM(nil, nil, 0)
	vm.SetHost(host)

	_, lineEnding := resolveConsoleOutputTarget(vm, consoleMethodFormats["log"])
	if lineEnding != "\n" {
		t.Fatalf("expected default line ending to be newline, got %q", lineEnding)
	}
}

// TestConsoleTimeAndTimeEnd verifies timer labels are stored and consumed by console.time/timeEnd.
func TestConsoleTimeAndTimeEnd(t *testing.T) {
	host := NewMockHost()
	var out bytes.Buffer
	host.SetOutput(&out)
	host.Request().ServerVars.Add("AXONASP_CLI_TUI", "1")

	vm := NewVM(nil, nil, 0)
	vm.SetHost(host)

	consoleDispatch(vm, "time", []Value{NewString("phase")})
	if _, exists := vm.consoleTimerItems["phase"]; !exists {
		t.Fatalf("expected timer label to be stored")
	}

	time.Sleep(2 * time.Millisecond)
	consoleDispatch(vm, "timeEnd", []Value{NewString("phase")})

	if _, exists := vm.consoleTimerItems["phase"]; exists {
		t.Fatalf("expected timer label to be removed after timeEnd")
	}

	rendered := out.String()
	if !strings.Contains(rendered, "phase:") || !strings.Contains(rendered, "ms") {
		t.Fatalf("expected timeEnd output with milliseconds, got %q", rendered)
	}
}

// TestConsoleDirFormatsObject verifies console.dir prints inspection-friendly JSON output.
func TestConsoleDirFormatsObject(t *testing.T) {
	host := NewMockHost()
	var out bytes.Buffer
	host.SetOutput(&out)
	host.Request().ServerVars.Add("AXONASP_CLI_TUI", "1")

	vm := NewVM(nil, nil, 0)
	vm.SetHost(host)

	objID := vm.allocJSID()
	vm.jsObjectItems[objID] = map[string]Value{
		"name": NewString("AxonASP"),
		"ok":   NewBool(true),
	}

	consoleDispatch(vm, "dir", []Value{{Type: VTJSObject, Num: objID}})
	rendered := out.String()
	if !strings.Contains(rendered, "\"name\": \"AxonASP\"") || !strings.Contains(rendered, "\"ok\": true") {
		t.Fatalf("expected dir output to include object properties, got %q", rendered)
	}
}

// TestConsoleTraceJScriptOnly verifies console.trace is emitted only for JScript runtime contexts.
func TestConsoleTraceJScriptOnly(t *testing.T) {
	host := NewMockHost()
	var out bytes.Buffer
	host.SetOutput(&out)
	host.Request().ServerVars.Add("AXONASP_CLI_TUI", "1")

	vm := NewVM(nil, nil, 0)
	vm.SetHost(host)

	fnID := vm.allocJSID()
	vm.jsFunctionItems[fnID] = &jsFunctionObject{name: "traceTarget"}
	vm.jsCallStack = append(vm.jsCallStack, jsCallFrame{
		fn:         Value{Type: VTJSFunction, Num: fnID},
		callLine:   41,
		callColumn: 7,
		callFile:   "www/tests/test_console_extensions.js",
	})
	vm.jsActiveEnvID = 1
	vm.lastLine = 52
	vm.lastColumn = 13
	vm.sourceName = "www/tests/test_console_extensions.js"

	consoleDispatch(vm, "trace", []Value{NewString("hello")})
	rendered := out.String()
	if !strings.Contains(rendered, "Trace: hello") || !strings.Contains(rendered, "traceTarget") || !strings.Contains(rendered, ":52:13") {
		t.Fatalf("expected trace output with frame metadata, got %q", rendered)
	}

	out.Reset()
	vm.jsCallStack = vm.jsCallStack[:0]
	vm.jsActiveEnvID = 0
	vm.jsRootEnvID = 0
	vm.engineMode = EngineModeVBScript
	consoleDispatch(vm, "trace", []Value{NewString("vb")})
	if out.Len() != 0 {
		t.Fatalf("expected no trace output for non-JScript contexts, got %q", out.String())
	}
}

// TestConsoleMethodsSerializeAllArgs verifies console methods include all provided arguments.
func TestConsoleMethodsSerializeAllArgs(t *testing.T) {
	host := NewMockHost()
	var out bytes.Buffer
	host.SetOutput(&out)
	host.Request().ServerVars.Add("AXONASP_CLI_TUI", "1")

	vm := NewVM(nil, nil, 0)
	vm.SetHost(host)

	methods := []struct {
		name     string
		expected string
	}{
		{name: "log", expected: "LOG"},
		{name: "info", expected: "INFO"},
		{name: "warn", expected: "WARN"},
		{name: "error", expected: "ERROR"},
		{name: "err", expected: "ERROR"},
	}

	for _, tc := range methods {
		out.Reset()
		consoleDispatch(vm, tc.name, []Value{NewString("Soma:"), NewInteger(8)})
		rendered := out.String()
		if !strings.Contains(rendered, "Soma: 8") {
			t.Fatalf("expected %s to include all args, got %q", tc.name, rendered)
		}
		if !strings.Contains(rendered, "["+tc.expected+"]") {
			t.Fatalf("expected %s output level %s, got %q", tc.name, tc.expected, rendered)
		}
	}
}
