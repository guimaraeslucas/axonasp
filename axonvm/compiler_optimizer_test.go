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
 * made available under the same license terms.
 */
package axonvm

import (
	"bytes"
	"encoding/binary"
	"testing"
)

// scanBytecodeForOp returns true if the given opcode appears anywhere in the
// bytecode stream after correct instruction-boundary parsing.
func scanBytecodeForOp(bytecode []byte, target OpCode) bool {
	for i := 0; i < len(bytecode); {
		op := OpCode(bytecode[i])
		if op == target {
			return true
		}
		i += 1 + opcodeOperandSize(op, bytecode, i)
	}
	return false
}

// scanBytecodeForExtOp returns true if the given extended opcode appears in the
// bytecode stream after instruction-boundary parsing.
func scanBytecodeForExtOp(bytecode []byte, target ExtOpCode) bool {
	for i := 0; i < len(bytecode); {
		op := OpCode(bytecode[i])
		if op == OpExtPrefix && i+1 < len(bytecode) {
			ext := ExtOpCode(bytecode[i+1])
			if ext == target {
				return true
			}
		}
		i += 1 + opcodeOperandSize(op, bytecode, i)
	}
	return false
}

// countBytecodeOp counts occurrences of the given opcode in bytecode.
func countBytecodeOp(bytecode []byte, target OpCode) int {
	n := 0
	for i := 0; i < len(bytecode); {
		op := OpCode(bytecode[i])
		if op == target {
			n++
		}
		i += 1 + opcodeOperandSize(op, bytecode, i)
	}
	return n
}

// runVBSAndGetOutput is a test helper that compiles an inline ASP source,
// runs it, and returns the buffered output string.
func runVBSAndGetOutput(t *testing.T, source string) string {
	t.Helper()
	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}
	vm := NewVMFromCompiler(compiler)
	host := NewMockHost()
	var buf bytes.Buffer
	host.SetOutput(&buf)
	vm.SetHost(host)
	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()
	return buf.String()
}

// TestVBScriptIntegerAddFolding verifies that "5 + 10" is folded to OpConstant(15)
// at compile time and OpAdd does NOT appear in the emitted bytecode.
func TestVBScriptIntegerAddFolding(t *testing.T) {
	source := `<% Dim x : x = 5 + 10 : Response.Write x %>`
	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	// The folded constant 15 must be present in the constant pool.
	found := false
	for _, c := range compiler.Constants() {
		if c.Type == VTInteger && c.Num == 15 {
			found = true
			break
		}
	}
	if !found {
		t.Error("constant pool does not contain folded integer 15")
	}

	// OpAdd must have been eliminated.
	if scanBytecodeForOp(compiler.Bytecode(), OpAdd) {
		t.Error("bytecode still contains OpAdd after constant folding")
	}

	// The program must still produce the correct output.
	out := runVBSAndGetOutput(t, source)
	if out != "15" {
		t.Errorf("unexpected output: got %q, want %q", out, "15")
	}
}

// TestVBScriptStringConcatFolding verifies that "hello" & " world" is folded
// to a single OpConstant("hello world") and OpConcat does not appear.
func TestVBScriptStringConcatFolding(t *testing.T) {
	source := `<% Dim s : s = "hello" & " world" : Response.Write s %>`
	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	found := false
	for _, c := range compiler.Constants() {
		if c.Type == VTString && c.Str == "hello world" {
			found = true
			break
		}
	}
	if !found {
		t.Error("constant pool does not contain folded string \"hello world\"")
	}
	if scanBytecodeForOp(compiler.Bytecode(), OpConcat) {
		t.Error("bytecode still contains OpConcat after constant folding")
	}

	out := runVBSAndGetOutput(t, source)
	if out != "hello world" {
		t.Errorf("unexpected output: got %q, want %q", out, "hello world")
	}
}

// TestVBScriptChainedFolding verifies that "a" & "b" & "c" collapses fully to
// the single string "abc" across multiple peephole passes.
func TestVBScriptChainedFolding(t *testing.T) {
	source := `<% Dim s : s = "a" & "b" & "c" : Response.Write s %>`
	out := runVBSAndGetOutput(t, source)
	if out != "abc" {
		t.Errorf("unexpected output for chained concat: got %q, want %q", out, "abc")
	}

	// Compile separately to inspect bytecode.
	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}
	if scanBytecodeForOp(compiler.Bytecode(), OpConcat) {
		t.Error("bytecode still contains OpConcat after chained constant folding")
	}
}

// TestVBScriptMulFolding verifies integer multiplication constant folding.
func TestVBScriptMulFolding(t *testing.T) {
	source := `<% Response.Write 1024 * 1024 %>`
	out := runVBSAndGetOutput(t, source)
	if out != "1048576" {
		t.Errorf("unexpected output: got %q, want %q", out, "1048576")
	}
	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}
	if scanBytecodeForOp(compiler.Bytecode(), OpMul) {
		t.Error("bytecode still contains OpMul after constant folding")
	}
}

// TestVBScriptJumpIntegrity verifies that constant-folded bytecode does not
// corrupt jump offsets: an If/Then/Else must execute the correct branch even
// when foldable constants appear in the condition and branches.
func TestVBScriptJumpIntegrity(t *testing.T) {
	// The condition "5 < 10" is NOT folded (comparison operators are not in the
	// fold list), so the jump opcode must still work correctly alongside the
	// folded constants in the branch bodies.
	source := `<% If 5 < 10 Then
	Response.Write "yes" & ""
Else
	Response.Write "no" & ""
End If %>`
	out := runVBSAndGetOutput(t, source)
	if out != "yes" {
		t.Errorf("jump integrity broken: got %q, want %q", out, "yes")
	}
}

// TestVBScriptJumpIntegrityElseBranch exercises the Else path.
func TestVBScriptJumpIntegrityElseBranch(t *testing.T) {
	source := `<% If 10 < 5 Then
	Response.Write "yes" & ""
Else
	Response.Write "no" & ""
End If %>`
	out := runVBSAndGetOutput(t, source)
	if out != "no" {
		t.Errorf("jump integrity (else) broken: got %q, want %q", out, "no")
	}
}

// TestVBScriptNopInBytecode verifies that OpNop placeholders are actually
// present after folding and that the VM executes past them without error.
func TestVBScriptNopInBytecode(t *testing.T) {
	source := `<% Response.Write(3 + 4) %>`
	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}
	// After folding "3 + 4" → 7, there should be at least one OpNop in the stream.
	if !scanBytecodeForOp(compiler.Bytecode(), OpNop) {
		t.Error("expected at least one OpNop in bytecode after folding, found none")
	}
	out := runVBSAndGetOutput(t, source)
	if out != "7" {
		t.Errorf("unexpected output: got %q, want %q", out, "7")
	}
}

// TestJScriptConstantFoldingBasic verifies JScript AST constant folding:
// "5 + 10" must produce 15 without emitting OpJSAdd.
func TestJScriptConstantFoldingBasic(t *testing.T) {
	source := `<%@ Language="JScript" %>
<% var x = 5 + 10; Response.Write(x); %>`
	out := runVBSAndGetOutput(t, source)
	if out != "15" {
		t.Errorf("JScript constant folding: got %q, want %q", out, "15")
	}
}

// TestJScriptStringFolding verifies that JScript string + string is folded.
func TestJScriptStringFolding(t *testing.T) {
	source := `<%@ Language="JScript" %>
<% var s = "foo" + "bar"; Response.Write(s); %>`
	out := runVBSAndGetOutput(t, source)
	if out != "foobar" {
		t.Errorf("JScript string folding: got %q, want %q", out, "foobar")
	}
}

// TestJScriptChainedFolding verifies that deeply nested constant arithmetic
// collapses at compile time.
func TestJScriptChainedFolding(t *testing.T) {
	source := `<%@ Language="JScript" %>
<% Response.Write(2 * 3 + 4); %>`
	out := runVBSAndGetOutput(t, source)
	if out != "10" {
		t.Errorf("JScript chained folding: got %q, want %q", out, "10")
	}
}

// TestVBScriptDirectMathOpcodes verifies that hot unary math builtins compile
// to dedicated opcodes instead of generic OpCall dispatch.
func TestVBScriptDirectMathOpcodes(t *testing.T) {
	source := `<% Response.Write Int(3.9) & "|" & Round(2.5) & "|" & Abs(-9) & "|" & Sqr(81) %>`
	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	if !scanBytecodeForOp(compiler.Bytecode(), OpMathInt) {
		t.Error("expected OpMathInt in bytecode")
	}
	if !scanBytecodeForOp(compiler.Bytecode(), OpMathRound) {
		t.Error("expected OpMathRound in bytecode")
	}
	if !scanBytecodeForOp(compiler.Bytecode(), OpMathAbs) {
		t.Error("expected OpMathAbs in bytecode")
	}
	if !scanBytecodeForOp(compiler.Bytecode(), OpMathSqr) {
		t.Error("expected OpMathSqr in bytecode")
	}

	out := runVBSAndGetOutput(t, source)
	if out != "3|2|9|9" {
		t.Fatalf("unexpected direct-math output: got %q, want %q", out, "3|2|9|9")
	}
}

// TestVBScriptIntegerInferenceOpRewrite verifies arithmetic over inferred integer
// locals is rewritten to integer VM opcodes.
func TestVBScriptIntegerInferenceOpRewrite(t *testing.T) {
	source := `<% Dim a, b, c : a = 9 : b = 3 : c = a + b * 2 - 1 : Response.Write c %>`
	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	if !scanBytecodeForOp(compiler.Bytecode(), OpIAdd) {
		t.Error("expected OpIAdd in bytecode")
	}
	if !scanBytecodeForOp(compiler.Bytecode(), OpIMul) {
		t.Error("expected OpIMul in bytecode")
	}
	if !scanBytecodeForOp(compiler.Bytecode(), OpISub) {
		t.Error("expected OpISub in bytecode")
	}

	out := runVBSAndGetOutput(t, source)
	if out != "14" {
		t.Fatalf("unexpected integer inference output: got %q, want %q", out, "14")
	}
}

// TestVBScriptIntegerDivPow2ToShift verifies integer division by a power-of-two
// constant is rewritten to the dedicated integer right-shift opcode.
func TestVBScriptIntegerDivPow2ToShift(t *testing.T) {
	source := `<% Dim x, y : x = 64 : y = x \ 4 : Response.Write y %>`
	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	if !scanBytecodeForOp(compiler.Bytecode(), OpIRightShift) {
		t.Error("expected OpIRightShift in bytecode")
	}

	out := runVBSAndGetOutput(t, source)
	if out != "16" {
		t.Fatalf("unexpected integer shift output: got %q, want %q", out, "16")
	}
}

func findOpcodeOffset(bytecode []byte, target OpCode) int {
	for i := 0; i < len(bytecode); {
		op := OpCode(bytecode[i])
		if op == target {
			return i
		}
		i += 1 + opcodeOperandSize(op, bytecode, i)
	}
	return -1
}

// TestVBScriptDeadIfFalseElimination verifies optimizer removes unreachable true branch
// bytes for compile-time false conditions.
func TestVBScriptDeadIfFalseElimination(t *testing.T) {
	source := `<%
If False Then
	Response.Write "A"
	Response.Write "B"
End If
Response.Write "Z"
%>`
	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}
	out := runVBSAndGetOutput(t, source)
	if out != "Z" {
		t.Fatalf("unexpected output: got %q, want %q", out, "Z")
	}

	jumpPos := findOpcodeOffset(compiler.Bytecode(), OpJumpIfFalse)
	if jumpPos < 0 {
		t.Fatalf("expected OpJumpIfFalse in bytecode")
	}
	target := int(binary.BigEndian.Uint32(compiler.Bytecode()[jumpPos+1 : jumpPos+5]))
	for i := jumpPos + 5; i < target; i++ {
		if OpCode(compiler.Bytecode()[i]) != OpNop {
			t.Fatalf("expected unreachable branch byte %d to be OpNop, got %v", i, OpCode(compiler.Bytecode()[i]))
		}
	}
}

// TestOptimizeRedundantContiguousLoadWithPop verifies redundant contiguous load pattern
// is eliminated by local-copy propagation pass.
func TestOptimizeRedundantContiguousLoadWithPop(t *testing.T) {
	compiler := NewCompiler("")
	compiler.bytecode = []byte{
		byte(OpGetLocal), 0, 1,
		byte(OpGetLocal), 0, 1,
		byte(OpPop),
		byte(OpHalt),
	}
	if !compiler.optimizeLocalCopyPropagationPass() {
		t.Fatalf("expected optimizeLocalCopyPropagationPass to report changes")
	}
	if OpCode(compiler.bytecode[3]) != OpNop || OpCode(compiler.bytecode[4]) != OpNop || OpCode(compiler.bytecode[5]) != OpNop {
		t.Fatalf("expected second OpGetLocal bytes to be NOP-filled")
	}
	if OpCode(compiler.bytecode[6]) != OpNop {
		t.Fatalf("expected trailing OpPop to be NOP-filled")
	}
}

// TestVBScriptCallMemberInlineCachePopulation verifies OpCallMember reserve bytes are
// populated by the VM after first run.
func TestVBScriptCallMemberInlineCachePopulation(t *testing.T) {
	source := `<%
Class CacheProbe
	Public Function Val()
		Val = 7
	End Function
End Class
Dim o, i
Set o = New CacheProbe
For i = 1 To 3
	Response.Write o.Val()
Next
%>`
	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	callPos := findOpcodeOffset(compiler.Bytecode(), OpCallMember)
	if callPos < 0 {
		t.Fatalf("expected OpCallMember in bytecode")
	}
	cachePos := callPos + 1 + 2 + 2
	if binary.BigEndian.Uint32(compiler.Bytecode()[cachePos:cachePos+4]) != 0 {
		t.Fatalf("expected cold cache slot to be zero before execution")
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var buf bytes.Buffer
	host.SetOutput(&buf)
	vm.SetHost(host)
	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	if buf.String() != "777" {
		t.Fatalf("unexpected output: got %q, want %q", buf.String(), "777")
	}
	if binary.BigEndian.Uint32(vm.bytecode[cachePos:cachePos+4]) == 0 {
		t.Fatalf("expected inline cache slot to be populated after execution")
	}
}
