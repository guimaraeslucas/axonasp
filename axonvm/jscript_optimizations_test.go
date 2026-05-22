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
 */
package axonvm

import (
	"testing"
)

func TestJScriptMathOptimizations(t *testing.T) {
	source := `<script runat="server" language="JScript">` +
		`Response.Write(Math.sin(0) + "|");` +
		`Response.Write(Math.cos(0) + "|");` +
		`Response.Write(Math.floor(1.5) + "|");` +
		`Response.Write(Math.min(10, 20) + "|");` +
		`Response.Write(Math.max(10, 20));` +
		`</script>`
	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	bytecode := compiler.Bytecode()
	ops := map[OpCode]bool{
		OpJSMathSin:   false,
		OpJSMathCos:   false,
		OpJSMathFloor: false,
		OpJSMathMin:   false,
		OpJSMathMax:   false,
	}
	for i := 0; i < len(bytecode); i++ {
		op := OpCode(bytecode[i])
		if _, exists := ops[op]; exists {
			ops[op] = true
		}
	}
	for op, found := range ops {
		if !found {
			t.Errorf("expected opcode %v in bytecode", op)
		}
	}

	out := runASPSourceForTest(t, source)
	expected := "0|1|1|10|20"
	if out != expected {
		t.Errorf("expected %q, got %q", expected, out)
	}
}

func TestJScriptBitwiseOptimization(t *testing.T) {
	source := `<script runat="server" language="JScript">` +
		`var x = 10;` +
		`Response.Write(Math.floor(x / 2) + "|");` +
		`Response.Write((x / 2) | 0);` +
		`</script>`
	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	bytecode := compiler.Bytecode()
	rightShiftCount := 0
	for i := 0; i < len(bytecode); i++ {
		if OpCode(bytecode[i]) == OpJSRightShift {
			rightShiftCount++
		}
	}
	if rightShiftCount != 2 {
		t.Errorf("expected 2 OpJSRightShift opcodes, got %d", rightShiftCount)
	}

	out := runASPSourceForTest(t, source)
	expected := "5|5"
	if out != expected {
		t.Errorf("expected %q, got %q", expected, out)
	}
}

func TestJScriptIntegerArithmeticSpecialization(t *testing.T) {
	source := `<script runat="server" language="JScript">` +
		`let i = 0 | 0;` +
		`let j = i + 1;` +
		`let k = j - 1;` +
		`Response.Write(k);` +
		`</script>`
	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	bytecode := compiler.Bytecode()
	hasAdd := false
	hasSub := false
	for i := 0; i < len(bytecode); i++ {
		op := OpCode(bytecode[i])
		if op == OpJSAdd {
			hasAdd = true
		}
		if op == OpJSSubtract {
			hasSub = true
		}
	}
	if !hasAdd {
		t.Errorf("expected OpJSAdd in bytecode")
	}
	if !hasSub {
		t.Errorf("expected OpJSSubtract in bytecode")
	}

	out := runASPSourceForTest(t, source)
	if out != "0" {
		t.Errorf("expected 0, got %q", out)
	}
}
