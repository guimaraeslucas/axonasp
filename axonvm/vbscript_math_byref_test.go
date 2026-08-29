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
	"strconv"
	"strings"
	"testing"
)

// TestVBScriptFastMathReadsVariableByValue verifies that the fast unary math
// opcodes (OpMathSqr, OpMathInt, OpMathAbs, ...) dereference a ByRef slot
// reference (VTArgRef) pushed by OpArgGlobalRef and read the actual variable
// value, not the global slot index. Regression for Sqr(i)/Int(Sqr(i)) returning
// Sqrt(slotIndex) instead of Sqrt(i).
func TestVBScriptFastMathReadsVariableByValue(t *testing.T) {
	source := `<% Dim i, x : i = 361 : x = -9 : Response.Write Int(Sqr(i)) & "|" & Sqr(i) & "|" & Abs(x) %>`
	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	// The fast path must be exercised: OpMathSqr/OpMathInt/OpMathAbs present.
	if !scanBytecodeForOp(compiler.Bytecode(), OpMathSqr) {
		t.Fatal("expected OpMathSqr in bytecode (fast unary math path)")
	}
	if !scanBytecodeForOp(compiler.Bytecode(), OpMathInt) {
		t.Fatal("expected OpMathInt in bytecode (fast unary math path)")
	}
	if !scanBytecodeForOp(compiler.Bytecode(), OpMathAbs) {
		t.Fatal("expected OpMathAbs in bytecode (fast unary math path)")
	}

	out := runVBSAndGetOutput(t, source)
	if out != "19|19|9" {
		t.Fatalf("fast math read wrong value: got %q, want %q", out, "19|19|9")
	}
}

// TestVBScriptFastMathVariableLoop verifies the fast math opcodes produce the
// correct per-iteration value when the operand is a loop variable passed ByRef.
func TestVBScriptFastMathVariableLoop(t *testing.T) {
	source := `<% Dim i, out
For i = 358 To 361
    out = out & Int(Sqr(i)) & ","
Next
Response.Write out %>`
	out := runVBSAndGetOutput(t, source)
	if out != "18,18,18,19," {
		t.Fatalf("loop fast-math output mismatch: got %q, want %q", out, "18,18,18,19,")
	}
}

// TestVBScriptPrimeBenchmarkCount verifies the prime-number benchmark up to
// 50,000 yields exactly 5,133 primes. When the fast Sqr math path read the ByRef
// slot index instead of the variable value, this count drifted to 9,025.
func TestVBScriptPrimeBenchmarkCount(t *testing.T) {
	source := `<% Dim i, j, isPrime, limit, primeCount
limit = 50000
primeCount = 0
For i = 2 To limit
    isPrime = True
    For j = 2 To Int(Sqr(i))
        If i Mod j = 0 Then
            isPrime = False
            Exit For
        End If
    Next
    If isPrime Then primeCount = primeCount + 1
Next
Response.Write primeCount %>`
	out := runVBSAndGetOutput(t, source)
	got := strings.TrimSpace(out)
	if got != strconv.Itoa(5133) {
		t.Fatalf("prime count mismatch: got %q, want %q", got, "5133")
	}
}
