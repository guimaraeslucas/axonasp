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

// Package axonvm — loop fast-path regression suite.
//
// These tests guard the OpForNextFastInt (VBScript) and OpJSJumpIfLessFast
// (JScript) super-instruction paths introduced for loop-overhead reduction.
// They verify:
//   - Correct bytecode emission (fast opcodes are actually generated).
//   - Correct execution semantics (results match the reference values).
//   - Edge cases: zero-iteration ranges, large counts, decrement loops,
//     non-unit steps (must fall back to slow path), global counter variables,
//     and break/continue inside fast-path loops.
package axonvm

import (
	"bytes"
	"testing"
)

// ---- VBScript OpForNextFastInt ----

// TestForNextFastIntEmitted verifies that the compiler emits OpForNextFastInt
// for a standard ascending For...Next loop with a local counter and default step.
func TestForNextFastIntEmitted(t *testing.T) {
	src := `<%
Sub RunLoop()
	Dim i, total
	total = 0
	For i = 1 To 5
		total = total + i
	Next
	Response.Write total
End Sub
Call RunLoop()
%>`
	comp := NewASPCompiler(src)
	if err := comp.Compile(); err != nil {
		t.Fatalf("compile: %v", err)
	}
	hasFast := false
	for _, b := range comp.Bytecode() {
		if OpCode(b) == OpForNextFastInt {
			hasFast = true
			break
		}
	}
	if !hasFast {
		t.Fatal("expected OpForNextFastInt in bytecode for local unit-step For loop")
	}
}

// TestForNextFastIntSlowPathForGlobal verifies that a For loop whose counter is a
// module-level (global) variable is NOT compiled with OpForNextFastInt.
func TestForNextFastIntSlowPathForGlobal(t *testing.T) {
	src := `<%
Dim g_i
For g_i = 1 To 3
	Response.Write g_i & " "
Next
%>`
	comp := NewASPCompiler(src)
	if err := comp.Compile(); err != nil {
		t.Fatalf("compile: %v", err)
	}
	for _, b := range comp.Bytecode() {
		if OpCode(b) == OpForNextFastInt {
			t.Fatal("OpForNextFastInt must NOT be emitted for a global loop counter")
		}
	}
}

// TestForNextFastGlobalIntEmitted verifies that the compiler emits OpForNextFastGlobalInt
// for a unit-step For...Next loop whose counter and limit both resolve to global slots.
func TestForNextFastGlobalIntEmitted(t *testing.T) {
	src := `<%
Dim g_i, g_sum, g_limit, g_count
g_sum = 0
g_limit = 3
For g_i = 1 To g_limit
	g_sum = g_sum + g_i
Next
g_count = 1
g_count = g_count + 1
g_count = g_count - 1
Response.Write g_sum & "|" & g_i & "|" & g_count
%>`
	comp := NewASPCompiler(src)
	if err := comp.Compile(); err != nil {
		t.Fatalf("compile: %v", err)
	}

	hasFastGlobal := false
	hasIncGlobal := false
	hasDecGlobal := false
	for _, b := range comp.Bytecode() {
		switch OpCode(b) {
		case OpForNextFastGlobalInt:
			hasFastGlobal = true
		case OpIncGlobalInt:
			hasIncGlobal = true
		case OpDecGlobalInt:
			hasDecGlobal = true
		}
	}

	if !hasFastGlobal {
		t.Fatal("expected OpForNextFastGlobalInt in bytecode for global unit-step For loop")
	}
	if !hasIncGlobal {
		t.Fatal("expected OpIncGlobalInt in bytecode for global increment assignment")
	}
	if !hasDecGlobal {
		t.Fatal("expected OpDecGlobalInt in bytecode for global decrement assignment")
	}
}

// TestForNextFastGlobalIntCorrectness verifies that the new global fast paths
// still execute with the expected VBScript semantics.
func TestForNextFastGlobalIntCorrectness(t *testing.T) {
	src := `<%
Dim g_i, g_sum, g_limit, g_count
g_sum = 0
g_limit = 3
For g_i = 1 To g_limit
	g_sum = g_sum + g_i
Next
g_count = 1
g_count = g_count + 1
g_count = g_count - 1
Response.Write g_sum & "|" & g_i & "|" & g_count
%>`
	runVBAndExpect(t, src, "6|4|1")
}

// TestForNextFastIntSlowPathForNonUnitStep verifies that a loop with Step 2
// is NOT compiled with OpForNextFastInt.
func TestForNextFastIntSlowPathForNonUnitStep(t *testing.T) {
	src := `<%
Sub RunLoop()
	Dim i
	For i = 0 To 10 Step 2
		Response.Write i & " "
	Next
End Sub
Call RunLoop()
%>`
	comp := NewASPCompiler(src)
	if err := comp.Compile(); err != nil {
		t.Fatalf("compile: %v", err)
	}
	for _, b := range comp.Bytecode() {
		if OpCode(b) == OpForNextFastInt {
			t.Fatal("OpForNextFastInt must NOT be emitted for Step != ±1")
		}
	}
}

// TestForNextFastIntCorrectnessAscending verifies that a fast-path ascending loop
// accumulates the correct sum.  For i = 1 To 100: sum = sum + i.  Expected: 5050.
func TestForNextFastIntCorrectnessAscending(t *testing.T) {
	src := `<%
Sub RunLoop()
	Dim i, s
	s = 0
	For i = 1 To 100
		s = s + i
	Next
	Response.Write s
End Sub
Call RunLoop()
%>`
	runVBAndExpect(t, src, "5050")
}

// TestForNextFastIntCorrectnessDescending verifies a fast-path descending loop
// (Step -1) by counting down from 5 to 1 and printing each value.
func TestForNextFastIntCorrectnessDescending(t *testing.T) {
	src := `<%
Sub RunLoop()
	Dim i
	For i = 5 To 1 Step -1
		Response.Write i & " "
	Next
End Sub
Call RunLoop()
%>`
	runVBAndExpect(t, src, "5 4 3 2 1 ")
}

// TestForNextFastIntZeroIterations verifies that a loop whose start > limit (for
// ascending) executes zero times and the counter keeps its initial value.
func TestForNextFastIntZeroIterations(t *testing.T) {
	src := `<%
Sub RunLoop()
	Dim i
	i = 99
	For i = 10 To 1
		Response.Write "BODY"
	Next
	Response.Write i
End Sub
Call RunLoop()
%>`
	// After a zero-iteration ascending loop the counter should equal its init value (10).
	runVBAndExpect(t, src, "10")
}

// TestForNextFastIntPostLoopCounterValue verifies that after a completed ascending
// loop the counter is limit+1 (standard VBScript behavior).
func TestForNextFastIntPostLoopCounterValue(t *testing.T) {
	src := `<%
Sub RunLoop()
	Dim i
	For i = 1 To 5
	Next
	Response.Write i
End Sub
Call RunLoop()
%>`
	runVBAndExpect(t, src, "6")
}

// TestForNextFastIntPostLoopCounterDescend verifies that after a completed
// descending loop the counter is limit-1.
func TestForNextFastIntPostLoopCounterDescend(t *testing.T) {
	src := `<%
Sub RunLoop()
	Dim i
	For i = 5 To 1 Step -1
	Next
	Response.Write i
End Sub
Call RunLoop()
%>`
	runVBAndExpect(t, src, "0")
}

// TestForNextFastIntDeadLoopElision verifies that a loop with an empty body and
// constant bounds is fully elided at compile time: no OpForNextFastInt should
// appear in the bytecode, only a direct constant assignment.
func TestForNextFastIntDeadLoopElision(t *testing.T) {
	src := `<%
Sub RunLoop()
	Dim i
	For i = 1 To 1000
	Next
	Response.Write i
End Sub
Call RunLoop()
%>`
	comp := NewASPCompiler(src)
	if err := comp.Compile(); err != nil {
		t.Fatalf("compile: %v", err)
	}
	// With constant bounds and an empty body the loop must be fully elided.
	for _, b := range comp.Bytecode() {
		if OpCode(b) == OpForNextFastInt {
			t.Fatal("dead loop with constant bounds should be elided — OpForNextFastInt must not appear")
		}
	}
	// Verify the elided result is semantically correct (post-loop counter = limit+1).
	runVBAndExpect(t, src, "1001")
}

// TestForNextFastIntDeadLoopElisionZeroRange verifies elision when init > limit
// (loop would never execute).
func TestForNextFastIntDeadLoopElisionZeroRange(t *testing.T) {
	src := `<%
Sub RunLoop()
	Dim i
	i = 42
	For i = 10 To 1
	Next
	Response.Write i
End Sub
Call RunLoop()
%>`
	// Loop never runs; counter should stay at its initial value (10, not 42).
	runVBAndExpect(t, src, "10")
}

// TestForNextFastIntLargeRange exercises the fast path with 5000 iterations to
// confirm there is no off-by-one and the final counter value is correct.
func TestForNextFastIntLargeRange(t *testing.T) {
	src := `<%
Sub RunLoop()
	Dim i, s
	s = 0
	For i = 1 To 5000
		s = s + 1
	Next
	Response.Write s & "," & i
End Sub
Call RunLoop()
%>`
	runVBAndExpect(t, src, "5000,5001")
}

// TestForNextFastIntBreakExits verifies that Exit For inside a fast-path loop
// correctly terminates the loop.
func TestForNextFastIntBreakExits(t *testing.T) {
	src := `<%
Sub RunLoop()
	Dim i
	For i = 1 To 100
		If i = 5 Then Exit For
	Next
	Response.Write i
End Sub
Call RunLoop()
%>`
	runVBAndExpect(t, src, "5")
}

// TestForNextFastIntNestedLoops verifies that nested fast-path loops produce the
// correct output, confirming that each loop's counter is independent.
func TestForNextFastIntNestedLoops(t *testing.T) {
	src := `<%
Sub RunLoop()
	Dim i, j, s
	s = 0
	For i = 1 To 3
		For j = 1 To 3
			s = s + 1
		Next
	Next
	Response.Write s
End Sub
Call RunLoop()
%>`
	runVBAndExpect(t, src, "9")
}

// TestForNextFastIntNonConstantBounds verifies that when the bounds are runtime
// variables (not compile-time constants) the loop still produces the correct result.
func TestForNextFastIntNonConstantBounds(t *testing.T) {
	src := `<%
Sub RunLoop(n)
	Dim i, s
	s = 0
	For i = 1 To n
		s = s + i
	Next
	Response.Write s
End Sub
Call RunLoop(10)
%>`
	runVBAndExpect(t, src, "55")
}

// ---- JScript OpJSJumpIfLessFast ----

// TestJSJumpIfLessFastEmitted verifies that the compiler emits OpJSJumpIfLessFast
// for the canonical ascending for-loop pattern `for (var i = 0; i < N; i++)`.
func TestJSJumpIfLessFastEmitted(t *testing.T) {
	src := `<%@ Language="JScript" %>
<%
var sum = 0;
for (var i = 0; i < 100; i++) {
	sum += i;
}
Response.Write(sum);
%>`
	comp := NewASPCompiler(src)
	if err := comp.Compile(); err != nil {
		t.Fatalf("compile: %v", err)
	}
	hasFast := false
	for _, b := range comp.Bytecode() {
		if OpCode(b) == OpJSJumpIfLessFast {
			hasFast = true
			break
		}
	}
	if !hasFast {
		t.Fatal("expected OpJSJumpIfLessFast in bytecode for `i < numericLiteral` test")
	}
}

// TestJSJumpIfLessFastSlowPathForNonLiteral verifies that when the limit is not a
// numeric literal the optimizer falls back to the generic test path.
func TestJSJumpIfLessFastSlowPathForNonLiteral(t *testing.T) {
	src := `<%@ Language="JScript" %>
<%
var n = 10;
var sum = 0;
for (var i = 0; i < n; i++) {
	sum += i;
}
Response.Write(sum);
%>`
	comp := NewASPCompiler(src)
	if err := comp.Compile(); err != nil {
		t.Fatalf("compile: %v", err)
	}
	for _, b := range comp.Bytecode() {
		if OpCode(b) == OpJSJumpIfLessFast {
			t.Fatal("OpJSJumpIfLessFast must NOT be emitted when the limit is a variable")
		}
	}
}

// TestJSJumpIfLessFastCorrectnessSum verifies the result of a fast-path ascending
// for-loop that accumulates a sum.  Sum of 0..99 = 4950.
func TestJSJumpIfLessFastCorrectnessSum(t *testing.T) {
	src := `<%@ Language="JScript" %>
<%
var sum = 0;
for (var i = 0; i < 100; i++) {
	sum += i;
}
Response.Write(sum);
%>`
	runVBAndExpect(t, src, "4950")
}

// TestJSJumpIfLessFastZeroIterations verifies that a loop with init >= limit
// executes zero times.
func TestJSJumpIfLessFastZeroIterations(t *testing.T) {
	src := `<%@ Language="JScript" %>
<%
var s = "";
for (var i = 10; i < 5; i++) {
	s += "BODY";
}
Response.Write(s === "" ? "ok" : "fail");
%>`
	runVBAndExpect(t, src, "ok")
}

// TestJSJumpIfLessFastLargeRange exercises the fast path with 5000 iterations.
func TestJSJumpIfLessFastLargeRange(t *testing.T) {
	src := `<%@ Language="JScript" %>
<%
var s = 0;
for (var i = 0; i < 5000; i++) {
	s++;
}
Response.Write(s);
%>`
	runVBAndExpect(t, src, "5000")
}

// TestJSJumpIfLessFastBreakExits verifies that break inside a fast-path loop
// correctly terminates the loop.
func TestJSJumpIfLessFastBreakExits(t *testing.T) {
	src := `<%@ Language="JScript" %>
<%
var i;
for (i = 0; i < 100; i++) {
	if (i === 42) break;
}
Response.Write(i);
%>`
	runVBAndExpect(t, src, "42")
}

// TestJSJumpIfLessFastNestedLoops verifies that nested fast-path JScript loops
// produce the correct result.
func TestJSJumpIfLessFastNestedLoops(t *testing.T) {
	src := `<%@ Language="JScript" %>
<%
var count = 0;
for (var i = 0; i < 4; i++) {
	for (var j = 0; j < 4; j++) {
		count++;
	}
}
Response.Write(count);
%>`
	runVBAndExpect(t, src, "16")
}

// TestJSJumpIfLessFastFinalCounterValue verifies that the loop variable holds the
// correct value after loop completion (should equal the limit).
func TestJSJumpIfLessFastFinalCounterValue(t *testing.T) {
	src := `<%@ Language="JScript" %>
<%
var i;
for (i = 0; i < 5; i++) {}
Response.Write(i);
%>`
	runVBAndExpect(t, src, "5")
}

// ---- Shared helpers ----

// runVBAndExpect compiles and executes an ASP source (VBScript or JScript) and
// asserts that the Response.Write output equals want.
func runVBAndExpect(t *testing.T, src, want string) {
	t.Helper()
	comp := NewASPCompiler(src)
	if err := comp.Compile(); err != nil {
		t.Fatalf("compile: %v", err)
	}
	vm := NewVM(comp.Bytecode(), comp.Constants(), comp.GlobalsCount())
	host := NewMockHost()
	var out bytes.Buffer
	host.SetOutput(&out)
	vm.SetHost(host)
	if err := vm.Run(); err != nil {
		t.Fatalf("run: %v", err)
	}
	got := out.String()
	if got != want {
		t.Errorf("output mismatch\n  got:  %q\n  want: %q", got, want)
	}
}
