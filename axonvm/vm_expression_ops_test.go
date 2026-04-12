/*
 * AxonASP Server
 * Copyright (C) 2026 G3pix Ltda. All rights reserved.
 *
 * Developed by Lucas GuimarÃ£es - G3pix Ltda
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

	"g3pix.com.br/axonasp/vbscript"
)

// TestVMCompileSupportsAllVBScriptLiterals verifies all lexer-produced VBScript literals compile and execute.
func TestVMCompileSupportsAllVBScriptLiterals(t *testing.T) {
	source := `<%= True %>|<%= False %>|<%= Null %>|<%= Empty %>|<%= Nothing %>|<%= &H0F %>|<%= &O17 %>|<%= #2026-03-13# %>`
	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}

	parts := strings.Split(output.String(), "|")
	if len(parts) != 8 {
		t.Fatalf("unexpected literal output: %q", output.String())
	}
	if parts[0] != "True" || parts[1] != "False" || parts[2] != "Null" || parts[3] != "" || parts[4] != "" || parts[5] != "15" || parts[6] != "15" {
		t.Fatalf("unexpected literal rendering: %q", output.String())
	}
	if !strings.Contains(parts[7], "2026-03-13") && !strings.Contains(parts[7], "3/13/2026") && !strings.Contains(parts[7], "13/03/2026") {
		t.Fatalf("unexpected date literal rendering: %q", parts[7])
	}
}

// TestVMSupportsUnaryAndBinaryOperators verifies arithmetic, comparison, and logical operators compile and execute.
func TestVMSupportsUnaryAndBinaryOperators(t *testing.T) {
	source := `<%= +5 %>|<%= -5 %>|<%= 2 + 3 %>|<%= 7 - 2 %>|<%= 3 * 4 %>|<%= 7 / 2 %>|<%= 7 \ 2 %>|<%= 7 Mod 3 %>|<%= 2 ^ 3 %>|<%= 2 < 3 %>|<%= 2 > 3 %>|<%= 2 <= 2 %>|<%= 3 >= 4 %>|<%= 2 = 2 %>|<%= 2 <> 3 %>|<%= Not True %>|<%= True And False %>|<%= True Or False %>|<%= True Xor False %>|<%= True Eqv True %>|<%= False Imp True %>`
	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}

	expected := "5|-5|5|5|12|3.5|3|1|8|True|False|True|False|True|True|False|False|True|True|True|True"
	if output.String() != expected {
		t.Fatalf("unexpected operator output:\nexpected: %q\nactual:   %q", expected, output.String())
	}
}

// TestVMSupportsBitwiseNumericOperators verifies numeric logical operators execute in bitwise mode.
func TestVMSupportsBitwiseNumericOperators(t *testing.T) {
	source := `<%= 6 And 3 %>|<%= 6 Or 3 %>|<%= 6 Xor 3 %>|<%= 6 Eqv 3 %>|<%= 6 Imp 3 %>|<%= Not 6 %>`
	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}

	expected := "2|7|5|-6|-5|-7"
	if output.String() != expected {
		t.Fatalf("unexpected numeric logical output:\nexpected: %q\nactual:   %q", expected, output.String())
	}
}

// TestVMNullPropagationInArithmetic verifies that Null propagates through arithmetic and comparison
// operations, while '&' coerces Null to an empty string as in Classic ASP/VBScript.
func TestVMNullPropagationInArithmetic(t *testing.T) {
	// Null propagates through +, -, *, /, =, Not, and unary negation.
	source := `<% x = Null %><%= x + 1 %>|<%= x - 1 %>|<%= x * 5 %>|<%= x / 2 %>|<%= x & "hi" %>|<%= x = 1 %>|<%= Not x %>|<%= -x %>`
	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}

	expected := "Null|Null|Null|Null|hi|Null|Null|Null"
	if output.String() != expected {
		t.Fatalf("unexpected null propagation output:\nexpected: %q\nactual:   %q", expected, output.String())
	}
}

// TestVMConcatNullCoercesToEmpty verifies sample-style string concatenation with Null keeps surrounding text.
func TestVMConcatNullCoercesToEmpty(t *testing.T) {
	source := `<% x = Null %><%= "non existant property:" & x & "(" & TypeName(x) & ")<br>" %>`
	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}

	expected := "non existant property:(Null)<br>"
	if output.String() != expected {
		t.Fatalf("unexpected null concat output:\nexpected: %q\nactual:   %q", expected, output.String())
	}
}

// TestVMNullLogicalShortCircuit verifies VBScript short-circuit rules for logical operators with Null.
// Null And False = False, Null Or True = True, Null Xor True = Null, Null Eqv Nothing = Null.
func TestVMNullLogicalShortCircuit(t *testing.T) {
	source := `<% x = Null %><%= x And False %>|<%= x And True %>|<%= x Or True %>|<%= x Or False %>|<%= x Xor True %>|<%= False Imp x %>|<%= x Imp True %>`
	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}

	// False And Null = False (short-circuit); Null And True = Null
	// Null Or True = True (short-circuit); Null Or False = Null
	// Null Xor True = Null; False Imp Null = True; Null Imp True = True
	expected := "False|Null|True|Null|Null|True|True"
	if output.String() != expected {
		t.Fatalf("unexpected null logical short-circuit output:\nexpected: %q\nactual:   %q", expected, output.String())
	}
}

// TestVMPlusOperatorClassicVBScriptSemantics verifies the overloaded Classic VBScript
// '+' behavior across String, Numeric, Empty, and Null Variant subtypes.
func TestVMPlusOperatorClassicVBScriptSemantics(t *testing.T) {
	source := `<%
Dim e, n
e = Empty
n = Null
%><%= "Alpha" + "Beta" %>|<%= 2 + "3" %>|<%= e + "Text" %>|<%= e + 7 %>|<%= n + 1 %>`
	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}

	expected := "AlphaBeta|5|Text|7|Null"
	if output.String() != expected {
		t.Fatalf("unexpected plus operator output:\nexpected: %q\nactual:   %q", expected, output.String())
	}
}

// TestVMPlusOperatorTypeMismatch verifies Numeric + non-numeric String raises VBScript Error 13.
func TestVMPlusOperatorTypeMismatch(t *testing.T) {
	source := `<%= 2 + "apple" %>`
	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	err := vm.Run()
	if err == nil {
		t.Fatalf("expected Type mismatch runtime error")
	}

	vme, ok := err.(*VMError)
	if !ok {
		t.Fatalf("expected VMError, got %T: %v", err, err)
	}
	if vme.Code != vbscript.TypeMismatch {
		t.Fatalf("expected TypeMismatch code, got %v", vme.Code)
	}
}

// TestCoerceInt64BankersRounding verifies that implicit float-to-integer coercion in integer
// arithmetic operations uses Banker's Rounding (round-half-to-even), matching VBScript behavior.
// VBScript rounds 0.5 -> 0, 1.5 -> 2, 2.5 -> 2, 3.5 -> 4 (round to nearest even).
func TestCoerceInt64BankersRounding(t *testing.T) {
	// Each expression uses integer division (\) which calls coerceInt64 on both operands,
	// but the simplest way to trigger coerceInt64 on a Double is via \ (integer division).
	// We use a known Double literal and verify the rounded result.
	cases := []struct {
		script string
		want   string
	}{
		{`<%= 0.5 \ 1 %>`, "0"},   // 0.5 -> round-to-even -> 0
		{`<%= 1.5 \ 1 %>`, "2"},   // 1.5 -> round-to-even -> 2
		{`<%= 2.5 \ 1 %>`, "2"},   // 2.5 -> round-to-even -> 2
		{`<%= 3.5 \ 1 %>`, "4"},   // 3.5 -> round-to-even -> 4
		{`<%= 4.5 \ 1 %>`, "4"},   // 4.5 -> round-to-even -> 4
		{`<%= -0.5 \ 1 %>`, "0"},  // -0.5 -> round-to-even -> 0
		{`<%= -1.5 \ 1 %>`, "-2"}, // -1.5 -> round-to-even -> -2
		{`<%= 2.3 \ 1 %>`, "2"},   // 2.3 -> truncate rounds down -> 2 (same in both)
		{`<%= 2.7 \ 1 %>`, "3"},   // 2.7 -> round-to-even rounds up -> 3
	}

	for _, tc := range cases {
		compiler := NewASPCompiler(tc.script)
		if err := compiler.Compile(); err != nil {
			t.Fatalf("compile failed for %q: %v", tc.script, err)
		}
		vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
		host := NewMockHost()
		var output bytes.Buffer
		host.SetOutput(&output)
		vm.SetHost(host)
		if err := vm.Run(); err != nil {
			t.Fatalf("run failed for %q: %v", tc.script, err)
		}
		if got := output.String(); got != tc.want {
			t.Errorf("script %q: got %q, want %q", tc.script, got, tc.want)
		}
	}
}
