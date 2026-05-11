/*
 * AxonASP Server
 * Copyright (C) 2026 G3pix Ltda. All rights reserved.
 *
 * Developed by Lucas Guimaraes - G3pix Ltda
 * Contact: https://g3pix.com.br
 * Project URL: https://g3pix.com.br/axonasp
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 */
package axonvm

import "testing"

func TestJScriptFunctionLocalSlotsEmitLocalOpcodes(t *testing.T) {
	source := `<script runat="server" language="JScript">` +
		`function f(){ var i = 0; i = i + 1; return i; }` +
		`Response.Write(f());` +
		`</script>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	hasGetLocal := false
	hasSetLocal := false
	for i := 0; i < len(compiler.Bytecode()); i++ {
		switch OpCode(compiler.Bytecode()[i]) {
		case OpJSGetLocal:
			hasGetLocal = true
		case OpJSSetLocal:
			hasSetLocal = true
		}
	}
	if !hasGetLocal || !hasSetLocal {
		t.Fatalf("expected local slot opcodes in bytecode, got %v", compiler.Bytecode())
	}

	out := runASPSourceForTest(t, source)
	if out != "1" {
		t.Fatalf("unexpected output: %q", out)
	}
}

func TestJScriptFunctionLocalLoopUsesIncLocalOpcode(t *testing.T) {
	source := `<script runat="server" language="JScript">` +
		`function f(){ var i = 0; for (i = 0; i < 4; i++) {} return i; }` +
		`Response.Write(f());` +
		`</script>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	hasIncLocal := false
	for i := 0; i < len(compiler.Bytecode()); i++ {
		if OpCode(compiler.Bytecode()[i]) == OpJSIncLocal {
			hasIncLocal = true
			break
		}
	}
	if !hasIncLocal {
		t.Fatalf("expected OpJSIncLocal in bytecode, got %v", compiler.Bytecode())
	}

	out := runASPSourceForTest(t, source)
	if out != "4" {
		t.Fatalf("unexpected output: %q", out)
	}
}

func TestJScriptLocalSlotRespectsLetShadowing(t *testing.T) {
	source := `<script runat="server" language="JScript">` +
		`function f(){ var x = 1; { let x = 2; x = 3; } return x; }` +
		`Response.Write(f());` +
		`</script>`

	out := runASPSourceForTest(t, source)
	if out != "1" {
		t.Fatalf("unexpected output: %q", out)
	}
}

func BenchmarkJScriptFunctionFastInt1M(b *testing.B) {
	source := `<%@ Language="JScript" %><%` +
		`function runLoop() {` +
		`  var sum = 0;` +
		`  for (let i = 0; i < 1000000; i++) { sum = sum + i; }` +
		`  return sum;` +
		`}` +
		`Response.Write(runLoop());` +
		`%>`
	benchmarkASPExecutionOnly(b, source)
}
