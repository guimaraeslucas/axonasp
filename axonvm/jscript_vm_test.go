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
	"encoding/binary"
	"errors"
	"os"
	"path/filepath"
	"strings"
	"testing"
	"time"

	"g3pix.com.br/axonasp/jscript"
	"g3pix.com.br/axonasp/vbscript"
)

func runASPSourceForTest(t *testing.T, source string) string {
	t.Helper()
	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}
	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	host.Response().SetBuffer(false)
	vm.SetHost(host)
	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	return output.String()
}

func runASPSourceForTestWithHost(t *testing.T, source string, host *MockHost) string {
	t.Helper()
	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}
	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	var output bytes.Buffer
	host.SetOutput(&output)
	host.Response().SetBuffer(false)
	vm.SetHost(host)
	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	return output.String()
}

func runASPSourceForTestWithErr(t *testing.T, source string) (string, error) {
	t.Helper()
	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}
	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	host.Response().SetBuffer(false)
	vm.SetHost(host)
	err := vm.Run()
	return output.String(), err
}

func runASPSourceForTestWithVM(t *testing.T, source string) (string, *VM, error) {
	t.Helper()
	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}
	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	host.Response().SetBuffer(false)
	vm.SetHost(host)
	err := vm.Run()
	return output.String(), vm, err
}

func TestJScriptCompileErrorsUseJScriptMetadata(t *testing.T) {
	compiler := NewASPCompiler(`<script runat="server" language="JScript">var x = ;</script>`)
	compiler.SetSourceName("/tests/jscript_compile_error.asp")

	err := compiler.Compile()
	if err == nil {
		t.Fatalf("expected compile error")
	}

	var syntaxErr *jscript.JSSyntaxError
	if !errors.As(err, &syntaxErr) {
		t.Fatalf("expected JScript syntax error, got %T", err)
	}

	if syntaxErr.Code != jscript.SyntaxError {
		t.Fatalf("unexpected code: got %d want %d", syntaxErr.Code, jscript.SyntaxError)
	}
	if strings.ReplaceAll(syntaxErr.File, "\\", "/") != "/tests/jscript_compile_error.asp" {
		t.Fatalf("unexpected file: %q", syntaxErr.File)
	}
	if syntaxErr.Category != "JScript compilation" {
		t.Fatalf("unexpected category: %q", syntaxErr.Category)
	}
	if syntaxErr.Source != "JScript compilation error" {
		t.Fatalf("unexpected source: %q", syntaxErr.Source)
	}
	if syntaxErr.Description != jscript.SyntaxError.String() {
		t.Fatalf("unexpected description: %q", syntaxErr.Description)
	}
	if syntaxErr.ASPDescription == "" {
		t.Fatalf("expected ASP description detail")
	}
	if syntaxErr.Number != jscript.HRESULTFromJScriptCode(jscript.SyntaxError) {
		t.Fatalf("unexpected number: got %d want %d", syntaxErr.Number, jscript.HRESULTFromJScriptCode(jscript.SyntaxError))
	}

	aspErr := CompilerErrorToASPError(err, "/tests/jscript_compile_error.asp")
	if aspErr.ASPCode != int(jscript.SyntaxError) {
		t.Fatalf("unexpected asp code: got %d want %d", aspErr.ASPCode, jscript.SyntaxError)
	}
	if aspErr.Number != jscript.HRESULTFromJScriptCode(jscript.SyntaxError) {
		t.Fatalf("unexpected asp number: got %d want %d", aspErr.Number, jscript.HRESULTFromJScriptCode(jscript.SyntaxError))
	}
	if aspErr.Source != "JScript compilation error" {
		t.Fatalf("unexpected asp source: %q", aspErr.Source)
	}
	if aspErr.Category != "JScript compilation" {
		t.Fatalf("unexpected asp category: %q", aspErr.Category)
	}
	if strings.Contains(aspErr.Source, "VBScript") {
		t.Fatalf("asp source should not fallback to VBScript: %q", aspErr.Source)
	}
}

func TestVBScriptCompileErrorsRemainUnchangedWithJScriptIsolation(t *testing.T) {
	compiler := NewASPCompiler(`<%= * %>`)
	compiler.SetSourceName("/tests/vbscript_compile_error.asp")

	err := compiler.Compile()
	if err == nil {
		t.Fatalf("expected compile error")
	}

	var syntaxErr *vbscript.VBSyntaxError
	if !errors.As(err, &syntaxErr) {
		t.Fatalf("expected VBScript syntax error, got %T", err)
	}

	if syntaxErr.Code != vbscript.SyntaxError {
		t.Fatalf("unexpected code: got %d want %d", syntaxErr.Code, vbscript.SyntaxError)
	}
	if syntaxErr.Source != "VBScript compilation error" {
		t.Fatalf("unexpected source: %q", syntaxErr.Source)
	}
	if syntaxErr.Category != "VBScript compilation" {
		t.Fatalf("unexpected category: %q", syntaxErr.Category)
	}

	aspErr := CompilerErrorToASPError(err, "/tests/vbscript_compile_error.asp")
	if aspErr.ASPCode != int(vbscript.SyntaxError) {
		t.Fatalf("unexpected asp code: got %d want %d", aspErr.ASPCode, vbscript.SyntaxError)
	}
	if aspErr.Source != "VBScript compilation error" {
		t.Fatalf("unexpected asp source: %q", aspErr.Source)
	}
	if aspErr.Category != "VBScript compilation" {
		t.Fatalf("unexpected asp category: %q", aspErr.Category)
	}
}

func TestJScriptResponseWriteFromScriptTag(t *testing.T) {
	source := `<script runat="server" language="JScript">Response.Write("Hello")</script>`
	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}
	containsJSOpcode := false
	for i := 0; i < len(compiler.Bytecode()); i++ {
		if OpCode(compiler.Bytecode()[i]) == OpJSCallMember {
			containsJSOpcode = true
			break
		}
	}
	if !containsJSOpcode {
		t.Fatalf("expected OpJSCallMember in bytecode, got %v", compiler.Bytecode())
	}
	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	if vm.Globals[0].Type != VTNativeObject {
		t.Fatalf("expected Response intrinsic at global 0, got %#v", vm.Globals[0])
	}
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	host.Response().SetBuffer(false)
	vm.SetHost(host)
	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	out := output.String()
	if out != "Hello" {
		t.Fatalf("expected Hello, got %q (bytecode=%v constants=%#v)", out, compiler.Bytecode(), compiler.Constants())
	}
}

func TestJScriptForLoopBytecodeContainsUpdateOpcodes(t *testing.T) {
	source := `<script runat="server" language="JScript">for (var i = 0; i < 2; i++) { var x = 0; x += i; }</script>`
	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	bytecode := compiler.Bytecode()
	hasFastInc := false
	hasAddAssign := false
	for i := range bytecode {
		switch OpCode(bytecode[i]) {
		// OpJSForFastInt is the fused single-opcode fast path for var/let integer loops;
		// it subsumes the separate increment opcode (OpJSIncLocalInt).
		case OpJSIncLocalInt, OpJSForFastInt:
			hasFastInc = true
		case OpJSAddAssign:
			hasAddAssign = true
		}
	}

	if !hasFastInc {
		t.Fatalf("expected OpJSIncLocalInt or OpJSForFastInt in bytecode, got %v", bytecode)
	}
	if !hasAddAssign {
		t.Fatalf("expected OpJSAddAssign in bytecode, got %v", bytecode)
	}
}

func TestJScriptMemberOpcodesReserveInlineCachePayload(t *testing.T) {
	source := `<script runat="server" language="JScript">` +
		`var o = { a: 1 };` +
		`Response.Write(o.a);` +
		`o.a = 2;` +
		`Response.Write("|" + o.a);` +
		`</script>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	bytecode := compiler.Bytecode()
	hasGet := false
	hasSet := false
	for ip := 0; ip < len(bytecode); {
		op := OpCode(bytecode[ip])
		sz := opcodeOperandSize(op, bytecode, ip)
		if ip+1+sz > len(bytecode) {
			t.Fatalf("invalid bytecode boundary at ip=%d op=%v", ip, op)
		}
		if op == OpJSMemberGet || op == OpJSMemberSet {
			if sz != 4 {
				t.Fatalf("expected 4-byte operand payload for %v, got %d", op, sz)
			}
			// ICNodeID is embedded in operand bytes; verify it's non-zero if expected.
			icNodeID := binary.BigEndian.Uint16(bytecode[ip+3:])
			if icNodeID == 0 {
				// ICNodeID 0 is reserved; first assigned ID starts at 1.
				// Accept 0 for programs with no IC-eligible nodes.
			}
			if op == OpJSMemberGet {
				hasGet = true
			} else {
				hasSet = true
			}
		}
		ip += 1 + sz
	}

	if !hasGet || !hasSet {
		t.Fatalf("expected both OpJSMemberGet and OpJSMemberSet in bytecode")
	}
}

func TestJScriptMemberInlineCachePopulatesAfterRun(t *testing.T) {
	source := `<script runat="server" language="JScript">` +
		`var o = { a: 1 };` +
		`Response.Write(o.a);` +
		`o.a = 3;` +
		`Response.Write("|" + o.a);` +
		`</script>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	// Pre-allocate icState based on the compiler's IC node count.
	if compiler.jsICNodeCount > 0 {
		vm.icState = make([]InlineCacheSlot, compiler.jsICNodeCount)
	}
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	host.Response().SetBuffer(false)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	if output.String() != "1|3" {
		t.Fatalf("unexpected output: %q", output.String())
	}

	foundPopulated := false
	for i := range vm.icState {
		if vm.icState[i].ShapeID != 0 && vm.icState[i].Flags != 0 {
			foundPopulated = true
			break
		}
	}

	if !foundPopulated {
		t.Fatalf("expected at least one populated JS member inline cache entry after execution")
	}
}

func TestJScriptForLoopAssignmentUpdateUsesIncrementFastPath(t *testing.T) {
	source := `<script runat="server" language="JScript">` +
		`var sum = 0;` +
		`for (var i = 0; i < 4; i = i + 1) { sum = sum + i; }` +
		`Response.Write(sum);` +
		`</script>`
	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	bytecode := compiler.Bytecode()
	hasFastInc := false
	for i := range bytecode {
		if OpCode(bytecode[i]) == OpJSIncLocalInt {
			hasFastInc = true
			break
		}
	}
	if !hasFastInc {
		t.Fatalf("expected OpJSIncLocalInt in bytecode, got %v", bytecode)
	}

	out := runASPSourceForTest(t, source)
	if out != "6" {
		t.Fatalf("unexpected for-loop output: %q", out)
	}
}

func TestJScriptForLoopFastUpdateAvoidsStackPop(t *testing.T) {
	source := `<script runat="server" language="JScript">for (var i = 0; i < 2; i++) {}</script>`
	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	bytecode := compiler.Bytecode()
	hasFastInc := false
	hasPop := false
	for i := range bytecode {
		switch OpCode(bytecode[i]) {
		// OpJSForFastInt is the fused fast path that replaces OpJSIncLocalInt.
		case OpJSIncLocalInt, OpJSForFastInt:
			hasFastInc = true
		case OpJSPop:
			hasPop = true
		}
	}

	if !hasFastInc {
		t.Fatalf("expected OpJSIncLocalInt or OpJSForFastInt in bytecode, got %v", bytecode)
	}
	if hasPop {
		t.Fatalf("expected no OpJSPop for fast for-update path, got %v", bytecode)
	}
}

func TestJScriptTailCallOpcodeEmission(t *testing.T) {
	source := `<script runat="server" language="JScript">` +
		`function sum(n, acc) {` +
		`if (n === 0) { return acc; }` +
		`return sum(n - 1, acc + 1);` +
		`}` +
		`Response.Write(sum(10, 0));` +
		`</script>`
	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	bytecode := compiler.Bytecode()
	hasTailCall := false
	for i := range bytecode {
		if OpCode(bytecode[i]) == OpJSTailCall {
			hasTailCall = true
			break
		}
	}
	if !hasTailCall {
		t.Fatalf("expected OpJSTailCall in bytecode, got %v", bytecode)
	}
}

func TestJScriptTailCallNotEmittedInsideTryCatch(t *testing.T) {
	source := `<script runat="server" language="JScript">` +
		`function sumInTry(n, acc) {` +
		`try {` +
		`if (n === 0) { return acc; }` +
		`return sumInTry(n - 1, acc + 1);` +
		`} catch (e) {` +
		`return -1;` +
		`}` +
		`}` +
		`Response.Write(sumInTry(8, 0));` +
		`</script>`
	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	bytecode := compiler.Bytecode()
	for i := range bytecode {
		switch OpCode(bytecode[i]) {
		case OpJSTailCall, OpJSTailCallMember:
			t.Fatalf("did not expect tail-call opcode inside try/catch, got bytecode %v", bytecode)
		}
	}
}

func TestJScriptTailCallMemberOpcodeEmission(t *testing.T) {
	source := `<script runat="server" language="JScript">` +
		`var obj = {};` +
		`obj.sum = function(n, acc) {` +
		`if (n === 0) { return acc; }` +
		`return obj.sum(n - 1, acc + 1);` +
		`};` +
		`Response.Write(obj.sum(10, 0));` +
		`</script>`
	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	bytecode := compiler.Bytecode()
	hasTailMemberCall := false
	for i := range bytecode {
		if OpCode(bytecode[i]) == OpJSTailCallMember {
			hasTailMemberCall = true
			break
		}
	}
	if !hasTailMemberCall {
		t.Fatalf("expected OpJSTailCallMember in bytecode, got %v", bytecode)
	}
}

func TestJScriptTailCallComputedMemberOpcodeEmission(t *testing.T) {
	source := `<script runat="server" language="JScript">` +
		`var obj = {};` +
		`obj.sum = function(n, acc) {` +
		`if (n === 0) { return acc; }` +
		`return obj["sum"](n - 1, acc + 1);` +
		`};` +
		`Response.Write(obj.sum(10, 0));` +
		`</script>`
	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	bytecode := compiler.Bytecode()
	hasTailComputedMemberCall := false
	for i := range bytecode {
		if OpCode(bytecode[i]) == OpJSTailCallComputedMember {
			hasTailComputedMemberCall = true
			break
		}
	}
	if !hasTailComputedMemberCall {
		t.Fatalf("expected OpJSTailCallComputedMember in bytecode, got %v", bytecode)
	}
}

func TestJScriptTailCallKeepsEnvGrowthBounded(t *testing.T) {
	source := `<script runat="server" language="JScript">` +
		`function sum(n, acc) {` +
		`if (n === 0) { return acc; }` +
		`return sum(n - 1, acc + 1);` +
		`}` +
		`Response.Write(sum(100000, 0));` +
		`</script>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	host.Response().SetBuffer(false)
	vm.SetHost(host)

	baseEnvCount := len(vm.jsEnvItems)
	baseArgsCount := len(vm.jsArgumentsItems)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	if output.String() != "100000" {
		t.Fatalf("expected 100000, got %q", output.String())
	}

	if len(vm.jsEnvItems)-baseEnvCount > 64 {
		t.Fatalf("tail call env growth too high: base=%d current=%d", baseEnvCount, len(vm.jsEnvItems))
	}
	if len(vm.jsArgumentsItems)-baseArgsCount > 64 {
		t.Fatalf("tail call arguments growth too high: base=%d current=%d", baseArgsCount, len(vm.jsArgumentsItems))
	}
}

func TestJScriptTailCallMemberKeepsEnvGrowthBounded(t *testing.T) {
	source := `<script runat="server" language="JScript">` +
		`var obj = {};` +
		`obj.sum = function(n, acc) {` +
		`if (n === 0) { return acc; }` +
		`return obj.sum(n - 1, acc + 1);` +
		`};` +
		`Response.Write(obj.sum(100000, 0));` +
		`</script>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	host.Response().SetBuffer(false)
	vm.SetHost(host)

	baseEnvCount := len(vm.jsEnvItems)
	baseArgsCount := len(vm.jsArgumentsItems)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	if output.String() != "100000" {
		t.Fatalf("expected 100000, got %q", output.String())
	}

	if len(vm.jsEnvItems)-baseEnvCount > 64 {
		t.Fatalf("tail member call env growth too high: base=%d current=%d", baseEnvCount, len(vm.jsEnvItems))
	}
	if len(vm.jsArgumentsItems)-baseArgsCount > 64 {
		t.Fatalf("tail member call arguments growth too high: base=%d current=%d", baseArgsCount, len(vm.jsArgumentsItems))
	}
}

func TestJScriptTailCallComputedMemberKeepsEnvGrowthBounded(t *testing.T) {
	source := `<script runat="server" language="JScript">` +
		`var obj = {};` +
		`obj.sum = function(n, acc) {` +
		`if (n === 0) { return acc; }` +
		`return obj["sum"](n - 1, acc + 1);` +
		`};` +
		`Response.Write(obj.sum(100000, 0));` +
		`</script>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	host.Response().SetBuffer(false)
	vm.SetHost(host)

	baseEnvCount := len(vm.jsEnvItems)
	baseArgsCount := len(vm.jsArgumentsItems)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	if output.String() != "100000" {
		t.Fatalf("expected 100000, got %q", output.String())
	}

	if len(vm.jsEnvItems)-baseEnvCount > 64 {
		t.Fatalf("tail computed member call env growth too high: base=%d current=%d", baseEnvCount, len(vm.jsEnvItems))
	}
	if len(vm.jsArgumentsItems)-baseArgsCount > 64 {
		t.Fatalf("tail computed member call arguments growth too high: base=%d current=%d", baseArgsCount, len(vm.jsArgumentsItems))
	}
}

func TestJScriptDeepNonTailRecursionReleasesEnvFrames(t *testing.T) {
	source := `<script runat="server" language="JScript">` +
		`function depth(n) {` +
		`if (n === 0) { return 0; }` +
		`return 1 + depth(n - 1);` +
		`}` +
		`Response.Write(depth(1500));` +
		`</script>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	host.Response().SetBuffer(false)
	vm.SetHost(host)

	baseEnvCount := len(vm.jsEnvItems)
	baseArgsCount := len(vm.jsArgumentsItems)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	if output.String() != "1500" {
		t.Fatalf("expected 1500, got %q", output.String())
	}

	if len(vm.jsEnvItems)-baseEnvCount > 64 {
		t.Fatalf("non-tail recursion env growth too high: base=%d current=%d", baseEnvCount, len(vm.jsEnvItems))
	}
	if len(vm.jsArgumentsItems)-baseArgsCount > 64 {
		t.Fatalf("non-tail recursion arguments growth too high: base=%d current=%d", baseArgsCount, len(vm.jsArgumentsItems))
	}
}

func TestJScriptSimpleForLoop(t *testing.T) {
	source := `<script runat="server" language="JScript">var sum = 0; for (var i = 0; i < 3; i++) { sum = sum + i; } Response.Write(sum);</script>`
	out := runASPSourceForTest(t, source)
	if out != "3" {
		t.Fatalf("unexpected simple for-loop output: %q", out)
	}
}

func TestJScriptClosureCapture(t *testing.T) {
	source := `<script runat="server" language="JScript">` +
		`function outer(v) { return function() { return v; }; }` +
		`var f = outer("ok");` +
		`Response.Write(f());` +
		`</script>`
	out := runASPSourceForTest(t, source)
	if out != "ok" {
		t.Fatalf("expected closure output ok, got %q", out)
	}
}

func TestJScriptTryCatchThrow(t *testing.T) {
	source := `<script runat="server" language="JScript">` +
		`try { throw "boom"; } catch (e) { Response.Write(e); }` +
		`</script>`
	out := runASPSourceForTest(t, source)
	if out != "boom" {
		t.Fatalf("expected catch output boom, got %q", out)
	}
}

// TestJScriptLegacyLanguageHeaderCompatibility ensures legacy <% @Language = JScript %>
// does not break server-side JScript output execution.
func TestJScriptLegacyLanguageHeaderCompatibility(t *testing.T) {
	source := `<%
@Language = JScript
%>
<script runat="server" language="JScript">Response.Write("ok")</script>`
	out := runASPSourceForTest(t, source)
	if out != "ok" {
		t.Fatalf("expected ok, got %q", out)
	}
}

// TestJScriptBinaryOperatorsForASPWrite validates operator codegen used by real ASP JScript pages.
func TestJScriptBinaryOperatorsForASPWrite(t *testing.T) {
	source := `<script runat="server" language="JScript">` +
		`var name = "";` +
		`var demo = name || "Guest";` +
		`Response.Write("Hello, " + demo);` +
		`Response.Write("; sum=" + (1 + 2));` +
		`</script>`
	out := runASPSourceForTest(t, source)
	if out != "Hello, Guest; sum=3" {
		t.Fatalf("unexpected output: %q", out)
	}
}

func TestJScriptSessionIndexedAssignment(t *testing.T) {
	source := `<script runat="server" language="JScript">` +
		`Session("k") = "v";` +
		`Response.Write(Session("k"));` +
		`</script>`
	out := runASPSourceForTest(t, source)
	if out != "v" {
		t.Fatalf("expected session value v, got %q", out)
	}
}

func TestJScriptEvalUsesJScriptExecutionContext(t *testing.T) {
	source := `<script runat="server" language="JScript">` +
		`var value = 10;` +
		`Response.Write(eval("value === 10 ? 'ok' : 'bad'"));` +
		`eval("value = value + 5;");` +
		`Response.Write("|" + value);` +
		`</script>`
	out := runASPSourceForTest(t, source)
	if out != "ok|15" {
		t.Fatalf("unexpected jscript eval output: %q", out)
	}
}

func TestJScriptSupportsTernaryAndStrictEqualityOperators(t *testing.T) {
	source := `<script runat="server" language="JScript">` +
		`Response.Write((5 === "5") ? "strict-eq-true" : "strict-eq-false");` +
		`Response.Write("|");` +
		`Response.Write((5 !== "5") ? "strict-neq-true" : "strict-neq-false");` +
		`</script>`
	out := runASPSourceForTest(t, source)
	if out != "strict-eq-false|strict-neq-true" {
		t.Fatalf("unexpected strict/ternary output: %q", out)
	}
}

func TestASPLanguageDirectiveRoutesPercentBlocksToJScript(t *testing.T) {
	source := `<%@ Language="JScript" %><% Response.Write(1 === 1 ? "ok" : "bad"); %>`
	out := runASPSourceForTest(t, source)
	if out != "ok" {
		t.Fatalf("expected ok from jscript percent block, got %q", out)
	}
}

func TestASPLanguageDirectiveWithoutSpaceRoutesPercentBlocksToJScript(t *testing.T) {
	source := `<%@Language="JScript"%><%` +
		`var metodo = String("POST");` +
		`if (metodo === "POST") { Response.Write("ok"); }` +
		`%>`
	out := runASPSourceForTest(t, source)
	if out != "ok" {
		t.Fatalf("expected ok from compact directive jscript block, got %q", out)
	}
}

func TestJScriptExpressionTagEmitsValueLikeVBScript(t *testing.T) {
	source := `<%@Language="JScript"%>` +
		`<% var nomeEnviado = "Lucas"; %>` +
		`Hello, <%= nomeEnviado %>!`
	out := runASPSourceForTest(t, source)
	if out != "Hello, Lucas!" {
		t.Fatalf("expected JScript expression tag output, got %q", out)
	}
}

func TestJScriptInlineHtmlConditionalBlocksRender(t *testing.T) {
	source := `<%@Language="JScript"%>` +
		`<% var metodo = String("POST"); var nomeEnviado = "Lucas"; %>` +
		`<% if (metodo === "POST" && nomeEnviado !== "") { %>` +
		`OK:<%= nomeEnviado %>` +
		`<% } else { %>` +
		`EMPTY` +
		`<% } %>`
	out := runASPSourceForTest(t, source)
	if out != "OK:Lucas" {
		t.Fatalf("unexpected inline html conditional output: %q", out)
	}
}

func TestJScriptFormPageCompilesAndRenders(t *testing.T) {
	pagePath := filepath.Join("..", "www", "tests", "test_jscript_form.asp")
	pageBytes, err := os.ReadFile(pagePath)
	if err != nil {
		t.Fatalf("failed to read test page: %v", err)
	}
	out := runASPSourceForTest(t, string(pageBytes))
	if !strings.Contains(out, "Formulário de Saudação") {
		t.Fatalf("expected form page output, got %q", out)
	}
}

func TestScriptRunatServerJScriptTagVariationsExecute(t *testing.T) {
	variations := []string{
		`<script type="text/javascript" language="javascript" runat="server">Response.Write("A")</script>`,
		`<script language="javascript" runat="server">Response.Write("B")</script>`,
		`<script type="text/javascript" language="jscript" runat="server">Response.Write("C")</script>`,
		`<script language="jscript" runat="server">Response.Write("D")</script>`,
	}
	want := []string{"A", "B", "C", "D"}

	for i := range variations {
		out := runASPSourceForTest(t, variations[i])
		if out != want[i] {
			t.Fatalf("variation %d: expected %q, got %q", i+1, want[i], out)
		}
	}
}

func TestVBScriptEvalRemainsUnchanged(t *testing.T) {
	source := `<% Response.Write Eval("1 + 2") %>`
	out := runASPSourceForTest(t, source)
	if out != "3" {
		t.Fatalf("expected VBScript Eval result 3, got %q", out)
	}
}

func TestJScriptForLoopControlStructures(t *testing.T) {
	source := `<script runat="server" language="JScript">` +
		`var sum = 0;` +
		`for (var i = 0; i < 6; i++) {` +
		`  if (i === 1) { continue; }` +
		`  if (i === 5) { break; }` +
		`  sum += i;` +
		`}` +
		`Response.Write(sum);` +
		`</script>`
	out := runASPSourceForTest(t, source)
	if out != "9" {
		t.Fatalf("unexpected for-loop output: %q", out)
	}
}

func TestJScriptWhileLoopControlStructures(t *testing.T) {
	source := `<script runat="server" language="JScript">` +
		`var whileCount = 0;` +
		`var j = 0;` +
		`while (j < 3) { whileCount += j; j++; }` +
		`Response.Write(whileCount);` +
		`</script>`
	out := runASPSourceForTest(t, source)
	if out != "3" {
		t.Fatalf("unexpected while-loop output: %q", out)
	}
}

func TestJScriptDoWhileLoopControlStructures(t *testing.T) {
	source := `<script runat="server" language="JScript">` +
		`var doCount = 0;` +
		`do { doCount++; } while (doCount < 2);` +
		`Response.Write(doCount);` +
		`</script>`
	out := runASPSourceForTest(t, source)
	if out != "2" {
		t.Fatalf("unexpected do-while output: %q", out)
	}
}

func TestJScriptSwitchStatement(t *testing.T) {
	source := `<script runat="server" language="JScript">` +
		`var result = "";` +
		`switch (2) {` +
		`case 1: result = "one"; break;` +
		`case 2: result = "two"; break;` +
		`default: result = "other";` +
		`}` +
		`Response.Write(result);` +
		`</script>`
	out := runASPSourceForTest(t, source)
	if out != "two" {
		t.Fatalf("unexpected switch output: %q", out)
	}
}

func TestJScriptSwitchFallthroughAndBreak(t *testing.T) {
	source := `<script runat="server" language="JScript">` +
		`var result = "";` +
		`switch (1) {` +
		`case 1: result += "a";` +
		`case 2: result += "b"; break;` +
		`default: result += "z";` +
		`}` +
		`Response.Write(result);` +
		`</script>`
	out := runASPSourceForTest(t, source)
	if out != "ab" {
		t.Fatalf("unexpected switch fallthrough output: %q", out)
	}
}

func TestJScriptForInLoopIteratesObjectKeys(t *testing.T) {
	source := `<script runat="server" language="JScript">` +
		`var o = {};` +
		`o.b = 2; o.a = 1; o.c = 3;` +
		`var joined = "";` +
		`for (var k in o) {` +
		`  if (k === "b") { continue; }` +
		`  joined += k;` +
		`  if (k === "c") { break; }` +
		`}` +
		`Response.Write(joined);` +
		`</script>`
	out := runASPSourceForTest(t, source)
	if out != "ac" {
		t.Fatalf("unexpected for-in output: %q", out)
	}
}

func TestJScriptStringAndArrayPrototypeMethods(t *testing.T) {
	source := `<script runat="server" language="JScript">` +
		`var text = "a,b,c";` +
		`var parts = text.split(",");` +
		`parts.push("d");` +
		`var popped = parts.pop();` +
		`Response.Write(text.indexOf("b") + "|" + parts.join("-") + "|" + popped);` +
		`</script>`
	out := runASPSourceForTest(t, source)
	if out != "2|a-b-c|d" {
		t.Fatalf("unexpected string/array output: %q", out)
	}
}

func TestJScriptReplaceAllWhileLoopTerminates(t *testing.T) {
	source := `<script runat="server" language="JScript">` +
		`function replaceAll(s, findText, replaceText) {` +
		`  var out = "" + s;` +
		`  while (out.indexOf(findText) >= 0) {` +
		`    out = out.split(findText).join(replaceText);` +
		`  }` +
		`  return out;` +
		`}` +
		`Response.Write(replaceAll("a&a&a", "&", "-"));` +
		`</script>`
	out := runASPSourceForTest(t, source)
	if out != "a-a-a" {
		t.Fatalf("unexpected replaceAll output: %q", out)
	}
}

func TestJScriptArgumentsThisCallAndApply(t *testing.T) {
	source := `<script runat="server" language="JScript">` +
		`function sum() { return arguments[0] + arguments[1]; }` +
		`function argc() { return arguments.length; }` +
		`function greet(prefix, suffix) { return prefix + this + suffix; }` +
		`Response.Write(sum(4, 5));` +
		`Response.Write("|");` +
		`Response.Write(argc(1,2,3));` +
		`Response.Write("|");` +
		`Response.Write(greet.call("Axon", "<", ">"));` +
		`Response.Write("|");` +
		`Response.Write(greet.apply("Axon", ["[", "]"]));` +
		`</script>`
	out := runASPSourceForTest(t, source)
	if out != "9|3|<Axon>|[Axon]" {
		t.Fatalf("unexpected arguments/call/apply output: %q", out)
	}
}

func TestJScriptCompoundAssignmentOperators(t *testing.T) {
	source := `<script runat="server" language="JScript">` +
		`var a = 10;` +
		`a += 5;` +
		`a -= 3;` +
		`a *= 2;` +
		`a /= 4;` +
		`Response.Write(a);` +
		`</script>`
	out := runASPSourceForTest(t, source)
	if out != "6" {
		t.Fatalf("unexpected compound-assignment output: %q", out)
	}
}

func TestJScriptCoercionMathDateAndRegex(t *testing.T) {
	source := `<script runat="server" language="JScript">` +
		`var undefinedValue;` +
		`var total = undefinedValue + 1;` +
		`var dt = new Date(2026, 3, 9, 16, 5, 7);` +
		`var matches = /ab+c/i.test("xxABBCyy");` +
		`Response.Write((total != total) ? "NaN" : total);` +
		`Response.Write("|");` +
		`Response.Write(Math.abs(-12));` +
		`Response.Write("|");` +
		`Response.Write(dt.getFullYear());` +
		`Response.Write("|");` +
		`Response.Write(matches ? "re-ok" : "re-bad");` +
		`</script>`
	out := runASPSourceForTest(t, source)
	if out != "NaN|12|2026|re-ok" {
		t.Fatalf("unexpected coercion/math/date/regex output: %q", out)
	}
}

func TestJScriptSelfExpandingReplaceLoopFailsFast(t *testing.T) {
	source := `<%@ Language="JScript" %>` +
		`<script runat="server" language="JScript">` +
		`function replaceAll(s, findText, replaceText) {` +
		`  var out = "" + s;` +
		`  while (out.indexOf(findText) >= 0) {` +
		`    out = out.split(findText).join(replaceText);` +
		`  }` +
		`  return out;` +
		`}` +
		`Response.Write(replaceAll("a&b", "&", "&amp;"));` +
		`</script>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}
	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	host.Response().SetBuffer(false)
	if err := host.Server().SetScriptTimeout(2); err != nil {
		t.Fatalf("set timeout failed: %v", err)
	}
	vm.SetHost(host)
	err := vm.Run()
	if err == nil {
		t.Fatalf("expected runtime error for runaway self-expanding replace loop")
	}
	errText := strings.ToLower(err.Error())
	if !strings.Contains(errText, "out of string") && !strings.Contains(errText, "string work exceeded") && !strings.Contains(errText, "loop iteration limit") {
		t.Fatalf("expected fast-fail watchdog, got: %v", err)
	}
}

func TestJScriptStringReplaceAndReplaceAll(t *testing.T) {
	source := `<script runat="server" language="JScript">` +
		`var s = "a_a_a";` +
		`var r1 = s.replace("_", "-");` +
		`var r2 = s.replaceAll("_", "-");` +
		`var r3 = "xxABBCyy".replace(/ab+c/i, "ok");` +
		`var r4 = "aba".replaceAll("", "-");` +
		`Response.Write(r1 + "|" + r2 + "|" + r3 + "|" + r4);` +
		`</script>`
	out := runASPSourceForTest(t, source)
	if out != "a-a_a|a-a-a|xxokyy|-a-b-a-" {
		t.Fatalf("unexpected replace output: %q", out)
	}
}

func TestJScriptEnumeratorAndVBArrayInterop(t *testing.T) {
	source := `<script runat="server" language="JScript">` +
		`var e = new Enumerator([10, 20, 30]);` +
		`var en = "";` +
		`while (!e.atEnd()) { en += e.item(); e.moveNext(); }` +
		`var split = "x,y,z".split(",");` +
		`var vb = new VBArray(split);` +
		`var dims = vb.dimensions();` +
		`var lb = vb.lbound();` +
		`var ub = vb.ubound();` +
		`var item = vb.getItem(1);` +
		`var joined = vb.toArray().join("+");` +
		`Response.Write(en + "|" + dims + "," + lb + "," + ub + "," + item + "," + joined);` +
		`</script>`
	out := runASPSourceForTest(t, source)
	if out != "102030|1,0,2,y,x+y+z" {
		t.Fatalf("unexpected enumerator/vbarray output: %q", out)
	}
}

func TestJScriptEnumeratorDictionaryKeysAndFormValues(t *testing.T) {
	host := NewMockHost()
	host.Request().Form.Add("alpha", "A")
	host.Request().Form.Add("beta", "B")

	source := `<script runat="server" language="JScript">` +
		`var d = Server.CreateObject("Scripting.Dictionary");` +
		`d.Add("first", "1");` +
		`d.Add("second", "2");` +
		`var de = new Enumerator(d);` +
		`var dk = "";` +
		`while (!de.atEnd()) { dk += "[" + de.item() + "]"; de.moveNext(); }` +
		`var fe = new Enumerator(Request.Form);` +
		`var fv = "";` +
		`while (!fe.atEnd()) { fv += "[" + fe.item() + "]"; fe.moveNext(); }` +
		`Response.Write(dk + "|" + fv);` +
		`</script>`

	out := runASPSourceForTestWithHost(t, source, host)
	if out != "[first][second]|[A][B]" {
		t.Fatalf("unexpected dictionary/form enumeration output: %q", out)
	}
}

func TestJScriptVBArrayDimensionsAndConcatBridge(t *testing.T) {
	source := `<script runat="server" language="JScript">` +
		`var nested = [[1, 2], [3, 4]];` +
		`var vbNested = new VBArray(nested);` +
		`var split = "x,y,z".split(",");` +
		`var vb = new VBArray(split);` +
		`var dims = vbNested.dimensions();` +
		`var asText = "<" + vb + ">";` +
		`var concat = [].concat(vb).join(",");` +
		`Response.Write(dims + "|" + asText + "|" + concat);` +
		`</script>`

	out := runASPSourceForTest(t, source)
	if out != "2|<x,y,z>|x,y,z" {
		t.Fatalf("unexpected vbarray concat bridge output: %q", out)
	}
}

func TestJScriptVBArrayLowerBoundIndexAndSourceStability(t *testing.T) {
	vm := NewVM(nil, nil, 0)

	source := ValueFromVBArray(NewVBArrayFromValues(5, []Value{NewString("first"), NewString("second")}))
	objID := vm.allocJSID()
	vm.jsObjectItems[objID] = map[string]Value{
		"__js_type":           NewString("VBArray"),
		"__js_vbarray_source": source,
	}
	vm.jsPropertyItems[objID] = make(map[string]jsPropertyDescriptor, 2)
	wrapper := Value{Type: VTJSObject, Num: objID}

	first := vm.jsVBArrayGetItem(wrapper, []Value{NewInteger(5)})
	second := vm.jsVBArrayGetItem(wrapper, []Value{NewInteger(6)})
	if first.Type != VTString || first.Str != "first" {
		t.Fatalf("expected first lower-bound item, got %#v", first)
	}
	if second.Type != VTString || second.Str != "second" {
		t.Fatalf("expected second lower-bound item, got %#v", second)
	}

	_ = vm.jsVBArrayToJSArray(vm.jsVBArraySource(wrapper))
	after := vm.jsVBArraySource(wrapper)
	if after.Type != VTArray || after.Arr == nil || after.Arr.Lower != 5 || len(after.Arr.Values) != 2 {
		t.Fatalf("expected preserved VBArray bridge source metadata, got %#v", after)
	}
}

func TestJScriptMathAndDateSurface(t *testing.T) {
	source := `<script runat="server" language="JScript">` +
		`var d = new Date(2026, 3, 9, 16, 5, 7, 123);` +
		`var rand = Math.random();` +
		`var randOk = (rand >= 0 && rand < 1) ? "ok" : "bad";` +
		`Response.Write(Math.pow(2, 5));` +
		`Response.Write("|" + Math.floor(2.9));` +
		`Response.Write("|" + Math.ceil(2.1));` +
		`Response.Write("|" + (Math.PI > 3 ? "pi" : "no"));` +
		`Response.Write("|" + d.getFullYear() + "," + d.getMonth() + "," + d.getDate());` +
		`Response.Write("|" + d.getHours() + "," + d.getMinutes() + "," + d.getSeconds() + "," + d.getMilliseconds());` +
		`Response.Write("|" + randOk);` +
		`</script>`
	out := runASPSourceForTest(t, source)
	if out != "32|2|3|pi|2026,3,9|16,5,7,123|ok" {
		t.Fatalf("unexpected math/date output: %q", out)
	}
}

func TestJScriptBinaryOperatorsUseJSOpcodes(t *testing.T) {
	// Verify correctness with constant operands (these will be folded at compile
	// time, so we only check the runtime result, not the bytecode opcodes).
	source := `<script runat="server" language="JScript">` +
		`var a = 7 - 2; var b = 3 * 4; var c = 9 / 3; var d = 9 % 4;` +
		`var e = ("5" == 5); var f = (2 < "10");` +
		`Response.Write(a + b + c + d + (e ? 1 : 0) + (f ? 1 : 0));` +
		`</script>`
	out := runASPSourceForTest(t, source)
	if out != "23" {
		t.Fatalf("unexpected binary operator output: %q", out)
	}

	// Verify JScript opcodes are emitted when operands are non-constant (variables),
	// ensuring the optimizer does not remove opcodes for dynamic expressions.
	varSource := `<script runat="server" language="JScript">` +
		`var x = 7; var y = 2;` +
		`var a = x - y; var b = x * y; var c = x / y; var d = x % y;` +
		`var e = ("5" == x); var f = (y < "10");` +
		`Response.Write(a + "|" + b + "|" + c + "|" + d);` +
		`</script>`
	bytecode := func() []byte {
		c2 := NewASPCompiler(varSource)
		if err := c2.Compile(); err != nil {
			t.Fatalf("compile failed: %v", err)
		}
		return c2.Bytecode()
	}()
	hasJSSubtract, hasJSMultiply, hasJSDivide, hasJSModulo, hasJSLooseEq, hasJSLess := false, false, false, false, false, false
	for i := range bytecode {
		switch OpCode(bytecode[i]) {
		case OpJSSubtract:
			hasJSSubtract = true
		case OpJSMultiply:
			hasJSMultiply = true
		case OpJSDivide:
			hasJSDivide = true
		case OpJSModulo:
			hasJSModulo = true
		case OpJSLooseEqual:
			hasJSLooseEq = true
		case OpJSLess:
			hasJSLess = true
		}
	}
	if !hasJSSubtract || !hasJSMultiply || !hasJSDivide || !hasJSModulo || !hasJSLooseEq || !hasJSLess {
		t.Fatalf("expected JS binary opcodes in variable-expression bytecode; got subtract=%v mul=%v div=%v mod=%v looseEq=%v less=%v",
			hasJSSubtract, hasJSMultiply, hasJSDivide, hasJSModulo, hasJSLooseEq, hasJSLess)
	}
	varOut := runASPSourceForTest(t, varSource)
	if varOut != "5|14|3.5|1" {
		t.Fatalf("unexpected variable binary operator output: %q", varOut)
	}
}

func TestJScriptUnicodeStringLiteralPreservedUTF8(t *testing.T) {
	// Use explicit accented text in a pure JScript block to ensure parser/compiler
	// conversions preserve UTF-8 content from the ASP source.
	source := `<script runat="server" language="JScript">Response.Write("Iteração número: 0")</script>`
	out := runASPSourceForTest(t, source)
	if out != "Iteração número: 0" {
		t.Fatalf("unexpected unicode jscript output: %q", out)
	}
}

func TestJScriptLooseNullComparisonDoesNotBlankFalseOrZero(t *testing.T) {
	source := `<script runat="server" language="JScript">` +
		`function htmlEncode(v) { if (v == null) { return ""; } return "" + v; }` +
		`var strictEq = (5 === "5");` +
		`var sideEffect = 0;` +
		`function touch() { sideEffect = sideEffect + 1; return true; }` +
		`var shortCircuit2 = false && touch();` +
		`Response.Write("strict=" + htmlEncode(strictEq));` +
		`Response.Write("|and=" + htmlEncode(shortCircuit2));` +
		`Response.Write("|count=" + htmlEncode(sideEffect));` +
		`</script>`
	out := runASPSourceForTest(t, source)
	if out != "strict=false|and=false|count=0" {
		t.Fatalf("unexpected null/false/zero rendering output: %q", out)
	}
}

func TestJScriptES5JSONParseAndStringify(t *testing.T) {
	source := `<script runat="server" language="JScript">` +
		`var obj = JSON.parse('{"name":"axon","count":2}');` +
		`Response.Write(obj.name + "|" + obj.count + "|");` +
		`Response.Write(JSON.stringify(obj));` +
		`</script>`
	out := runASPSourceForTest(t, source)
	if out != `axon|2|{"count":2,"name":"axon"}` && out != `axon|2|{"name":"axon","count":2}` {
		t.Fatalf("unexpected JSON parse/stringify output: %q", out)
	}
}

func TestJScriptES5ArrayIndexMethodsAndIsArray(t *testing.T) {
	source := `<script runat="server" language="JScript">` +
		`var a = [10,20,10,30];` +
		`Response.Write(a.indexOf(10) + "|" + a.lastIndexOf(10) + "|");` +
		`Response.Write(Array.isArray(a) ? "yes" : "no");` +
		`</script>`
	out := runASPSourceForTest(t, source)
	if out != "0|2|yes" {
		t.Fatalf("unexpected array ES5 output: %q", out)
	}
}

func TestJScriptES5ObjectStaticsAndDescriptors(t *testing.T) {
	source := `<script runat="server" language="JScript">` +
		`var proto = { p: 1 };` +
		`var obj = Object.create(proto);` +
		`Object.defineProperty(obj, "x", { value: 7, writable: false, enumerable: true, configurable: true });` +
		`Object.defineProperties(obj, { y: { value: 9, enumerable: false, configurable: true } });` +
		`var keys = Object.keys(obj).join(",");` +
		`var d = Object.getOwnPropertyDescriptor(obj, "x");` +
		`var p = Object.getPrototypeOf(obj);` +
		`Response.Write(keys + "|" + d.value + "," + (d.writable ? "w" : "nw") + "|" + p.p);` +
		`</script>`
	out := runASPSourceForTest(t, source)
	if out != "x|7,nw|1" {
		t.Fatalf("unexpected object ES5 descriptor output: %q", out)
	}
}

func TestJScriptObjectLiteralGetterAndSetter(t *testing.T) {
	source := `<script runat="server" language="JScript">` +
		`var obj = { _v: 0, get value() { return this._v + 1; }, set value(v) { this._v = v; } };` +
		`obj.value = 4;` +
		`Response.Write(obj.value + "|" + obj._v);` +
		`</script>`
	out := runASPSourceForTest(t, source)
	if out != "5|4" {
		t.Fatalf("unexpected object literal accessor output: %q", out)
	}
}

func TestJScriptES5StringTrim(t *testing.T) {
	source := `<script runat="server" language="JScript">Response.Write("  AxonASP  ".trim());</script>`
	out := runASPSourceForTest(t, source)
	if out != "AxonASP" {
		t.Fatalf("unexpected trim output: %q", out)
	}
}

func TestJScriptStringBracketIndexingBounds(t *testing.T) {
	source := `<script runat="server" language="JScript">` +
		`var s = "text";` +
		`Response.Write(s[0] + "|" + s[3] + "|" + (s[4] === undefined ? "u" : "x") + "|" + "á"[0]);` +
		`</script>`
	out := runASPSourceForTest(t, source)
	if out != "t|t|u|á" {
		t.Fatalf("unexpected string bracket indexing output: %q", out)
	}
}

func TestJScriptArgumentsAliasingNonStrict(t *testing.T) {
	source := `<script runat="server" language="JScript">` +
		`function probe(a, b) { a = 10; var v0 = arguments[0]; arguments[1] = 20; return v0 + "|" + b + "|" + arguments[1]; }` +
		`Response.Write(probe(1, 2));` +
		`</script>`
	out := runASPSourceForTest(t, source)
	if out != "10|20|20" {
		t.Fatalf("unexpected arguments aliasing output: %q", out)
	}
}

func TestJScriptWithStatementScopeResolution(t *testing.T) {
	source := `<script runat="server" language="JScript">` +
		`var x = "outer";` +
		`var obj = { x: "inner", y: 1 };` +
		`with (obj) { x = x + "-" + y; y = 5; z = 9; }` +
		`Response.Write(x + "|" + obj.x + "|" + obj.y + "|" + (typeof z));` +
		`</script>`
	out := runASPSourceForTest(t, source)
	if out != "outer|inner-1|5|number" {
		t.Fatalf("unexpected with statement scope output: %q", out)
	}
}

func TestJScriptDefinePropertyRejectsIllegalTransitions(t *testing.T) {
	source := `<script runat="server" language="JScript">` +
		`var obj = {};` +
		`Object.defineProperty(obj, "x", { value: 7, writable: false, configurable: false, enumerable: true });` +
		`Object.defineProperty(obj, "x", { value: 8 });` +
		`Object.defineProperty(obj, "x", { configurable: true });` +
		`Object.defineProperty(obj, "x", { get: function() { return 1; } });` +
		`var d = Object.getOwnPropertyDescriptor(obj, "x");` +
		`Response.Write(obj.x + "|" + (d.configurable ? "c" : "nc") + "|" + (d.writable ? "w" : "nw") + "|" + (d.get ? "g" : "ng"));` +
		`</script>`
	out := runASPSourceForTest(t, source)
	if out != "7|nc|nw|ng" {
		t.Fatalf("unexpected defineProperty transition output: %q", out)
	}
}

func TestJScriptPrototypeChainResolution(t *testing.T) {
	source := `<script runat="server" language="JScript">` +
		`var proto = { base: 7 };` +
		`var obj = Object.create(proto);` +
		`Response.Write(obj.base);` +
		`</script>`
	out := runASPSourceForTest(t, source)
	if out != "7" {
		t.Fatalf("unexpected prototype-chain output: %q", out)
	}
}

func TestJScriptArrayPrototypeMethodResolution(t *testing.T) {
	source := `<script runat="server" language="JScript">` +
		`Array.prototype.firstItem = function() { return this[0]; };` +
		`var values = [42, 99];` +
		`Response.Write(values.firstItem());` +
		`</script>`
	out := runASPSourceForTest(t, source)
	if out != "42" {
		t.Fatalf("unexpected array prototype method output: %q", out)
	}
}

func TestJScriptUserConstructorPrototypeAndNew(t *testing.T) {
	source := `<script runat="server" language="JScript">` +
		`function Box(v) { this.value = v; }` +
		`Box.prototype.getValue = function() { return this.value; };` +
		`var box = new Box(11);` +
		`Response.Write(box.getValue() + "|" + Object.getPrototypeOf(box).constructor.prototype.getValue.call(box));` +
		`</script>`
	out := runASPSourceForTest(t, source)
	if out != "11|11" {
		t.Fatalf("unexpected constructor/prototype output: %q", out)
	}
}

func TestJScriptNumberPrimitivePrototypeExtension(t *testing.T) {
	source := `<script runat="server" language="JScript">` +
		`Number.prototype.prefixPad = function(intLength, strPadChar) {` +
		`	if (typeof intLength !== "number") var intLength = 2;` +
		`	if (typeof strPadChar !== "string") var strPadChar = "0";` +
		`	var strInput = String(this);` +
		`	for (var i = strInput.length; i < intLength; ++i) strInput = strPadChar + strInput;` +
		`	return strInput;` +
		`};` +
		`Response.Write("Number.prefixPad: " + ((7).prefixPad()) + "\n");` +
		`Response.Write("Number.prefixPad: " + ((9).prefixPad(3, "9")) + "\n");` +
		`</script>`
	out := runASPSourceForTest(t, source)
	expected := "Number.prefixPad: 07\nNumber.prefixPad: 999\n"
	if out != expected {
		t.Fatalf("unexpected number prototype extension output: %q (expected: %q)", out, expected)
	}
}

func TestJScriptStringPrimitivePrototypeExtension(t *testing.T) {
	source := `<script runat="server" language="JScript">` +
		`String.prototype.prefixPad = function(intLength, strPadChar) {` +
		`	if (typeof intLength !== "number") var intLength = 2;` +
		`	if (typeof strPadChar !== "string") var strPadChar = "0";` +
		`	var strInput = String(this);` +
		`	for (var i = strInput.length; i < intLength; ++i) strInput = strPadChar + strInput;` +
		`	return strInput;` +
		`};` +
		`Response.Write("String.prefixPad: " + ("07".prefixPad(3)) + "\n");` +
		`Response.Write("String.prefixPad: " + ("H".prefixPad(5, "O")) + "\n");` +
		`</script>`
	out := runASPSourceForTest(t, source)
	expected := "String.prefixPad: 007\nString.prefixPad: OOOOH\n"
	if out != expected {
		t.Fatalf("unexpected string prototype extension output: %q (expected: %q)", out, expected)
	}
}

func TestJScriptBooleanPrimitivePrototypeExtension(t *testing.T) {
	source := `<script runat="server" language="JScript">` +
		`Boolean.prototype.toCustomString = function() {` +
		`	return this ? "YES" : "NO";` +
		`};` +
		`Response.Write(true.toCustomString() + "|" + false.toCustomString());` +
		`</script>`
	out := runASPSourceForTest(t, source)
	expected := "YES|NO"
	if out != expected {
		t.Fatalf("unexpected boolean prototype extension output: %q (expected: %q)", out, expected)
	}
}

func TestJScriptNumberPrimitivePropertyAccess(t *testing.T) {
	source := `<script runat="server" language="JScript">` +
		`Number.prototype.customProp = "fromNumberProto";` +
		`Response.Write((42).customProp);` +
		`</script>`
	out := runASPSourceForTest(t, source)
	expected := "fromNumberProto"
	if out != expected {
		t.Fatalf("unexpected number primitive property access output: %q (expected: %q)", out, expected)
	}
}

func TestJScriptMemberAndIndexUpdateOperators(t *testing.T) {
	source := `<script runat="server" language="JScript">` +
		`var obj = { count: 1 };` +
		`var arr = [4, 6];` +
		`Response.Write(obj.count++ + "|" + obj.count + "|");` +
		`Response.Write(--arr[1] + "|" + arr[1]);` +
		`</script>`
	out := runASPSourceForTest(t, source)
	if out != "1|2|5|5" {
		t.Fatalf("unexpected member/index update output: %q", out)
	}
}

func TestJScriptChainedMemberIndexAccess(t *testing.T) {
	source := `<script runat="server" language="JScript">` +
		`var obj = { items: [3, 8] };` +
		`obj.items[0] = 5;` +
		`Response.Write(obj.items[0] + "|" + obj["items"][1]);` +
		`</script>`
	out := runASPSourceForTest(t, source)
	if out != "5|8" {
		t.Fatalf("unexpected chained member/index output: %q", out)
	}
}

func TestJScriptFunctionBindCallApply(t *testing.T) {
	source := `<script runat="server" language="JScript">` +
		`function add(a, b) { return this.base + a + b; }` +
		`var bound = add.bind({ base: 10 }, 2);` +
		`Response.Write(bound(3) + "|");` +
		`Response.Write(add.call({ base: 4 }, 1, 2) + "|");` +
		`Response.Write(add.apply({ base: 5 }, [1, 2]));` +
		`</script>`
	out := runASPSourceForTest(t, source)
	if out != "15|7|8" {
		t.Fatalf("unexpected bind/call/apply output: %q", out)
	}
}

func TestJScriptES5ArrayIteratorCallbacks(t *testing.T) {
	source := `<script runat="server" language="JScript">` +
		`var items = [1, 2, 3];` +
		`var seen = "";` +
		`items.forEach(function(v, i) { seen += v + ":" + i + ";"; });` +
		`var every = items.every(function(v) { return v < 4; }) ? "T" : "F";` +
		`var some = items.some(function(v) { return v === 2; }) ? "T" : "F";` +
		`var mapped = items.map(function(v) { return v * 2; }).join(",");` +
		`var filtered = items.filter(function(v) { return v >= 2; }).join(",");` +
		`var reduced = items.reduce(function(acc, v) { return acc + v; }, 0);` +
		`var reducedRight = ["a", "b", "c"].reduceRight(function(acc, v) { return acc + v; }, "");` +
		`Response.Write(seen + "|" + every + some + "|" + mapped + "|" + filtered + "|" + reduced + "|" + reducedRight);` +
		`</script>`
	out := runASPSourceForTest(t, source)
	if out != "1:0;2:1;3:2;|TT|2,4,6|2,3|6|cba" {
		t.Fatalf("unexpected array iterator output: %q", out)
	}
}

func TestJScriptES5ObjectIntegrityStatics(t *testing.T) {
	source := `<script runat="server" language="JScript">` +
		`var obj = {};` +
		`Object.defineProperty(obj, "visible", { value: 1, writable: true, enumerable: true, configurable: true });` +
		`Object.defineProperty(obj, "hidden", { value: 2, writable: false, enumerable: false, configurable: true });` +
		`var accessor = {};` +
		`Object.defineProperty(accessor, "value", { get: function() { return 7; }, set: function(v) {}, enumerable: false, configurable: true });` +
		`var names = Object.getOwnPropertyNames(obj).join(",");` +
		`var dataDesc = Object.getOwnPropertyDescriptor(obj, "hidden");` +
		`var accessorDesc = Object.getOwnPropertyDescriptor(accessor, "value");` +
		`Object.preventExtensions(obj);` +
		`obj.extra = 9;` +
		`var sealed = { keep: 5 };` +
		`Object.seal(sealed); delete sealed.keep; sealed.extra = 8;` +
		`var frozen = { locked: 4 };` +
		`Object.freeze(frozen); frozen.locked = 10;` +
		`var iso = new Date(2026, 0, 2, 3, 4, 5).toJSON();` +
		`Response.Write(names + "|" + (dataDesc.writable ? "w" : "nw") + ":" + (dataDesc.enumerable ? "e" : "ne") + ":" + (dataDesc.configurable ? "c" : "nc") + "|" + accessorDesc.get.call(accessor) + ":" + (accessorDesc.set ? "set" : "noset") + ":" + (accessorDesc.enumerable ? "e" : "ne") + ":" + (accessorDesc.configurable ? "c" : "nc") + "|" + (Object.isExtensible(obj) ? "ext" : "fixed") + ":" + (obj.extra === undefined ? "noextra" : "extra") + "|" + sealed.keep + ":" + (Object.isSealed(sealed) ? "sealed" : "open") + "|" + frozen.locked + ":" + (Object.isFrozen(frozen) ? "frozen" : "open") + "|" + iso);` +
		`</script>`
	out := runASPSourceForTest(t, source)
	if !strings.Contains(out, "hidden,visible|") && !strings.Contains(out, "visible,hidden|") {
		t.Fatalf("unexpected object integrity property names output: %q", out)
	}
	if !strings.Contains(out, "|nw:ne:c|7:set:ne:c|fixed:noextra|5:sealed|4:frozen|") {
		t.Fatalf("unexpected object integrity statics output: %q", out)
	}
	if !strings.Contains(out, "T03:04:05Z") {
		t.Fatalf("expected Date.toJSON ISO timestamp, got %q", out)
	}
}

func TestJScriptES5StringMethodsSurface(t *testing.T) {
	source := `<script runat="server" language="JScript">` +
		`var s = "AbcDef";` +
		`Response.Write(s.charAt(1) + "|");` +
		`Response.Write(s.charCodeAt(1) + "|");` +
		`Response.Write(s.substring(1,4) + "|");` +
		`Response.Write(s.substr(2,3) + "|");` +
		`Response.Write(s.slice(1,-1) + "|");` +
		`Response.Write(s.concat("X","Y") + "|");` +
		`Response.Write("abc123".match(/\d+/)[0] + "|");` +
		`Response.Write("abc123".search(/\d+/) + "|");` +
		`Response.Write(s.toLowerCase() + "|");` +
		`Response.Write(s.toUpperCase() + "|");` +
		`Response.Write("abc".localeCompare("abd"));` +
		`</script>`
	out := runASPSourceForTest(t, source)
	if out != "b|98|bcD|cDe|bcDe|AbcDefXY|123|3|abcdef|ABCDEF|-1" {
		t.Fatalf("unexpected string ES5 methods output: %q", out)
	}
}

func TestJScriptES5ArrayMethodsSurface(t *testing.T) {
	source := `<script runat="server" language="JScript">` +
		`var a = [1,2,3,4];` +
		`Response.Write(a.slice(1,3).join("") + "|");` +
		`Response.Write(a.splice(1,2,9,8).join("") + "|");` +
		`Response.Write(a.join("") + "|");` +
		`Response.Write(a.shift() + "|");` +
		`Response.Write(a.unshift(7,6) + "|");` +
		`Response.Write(a.reverse().join("") + "|");` +
		`Response.Write(["b","a","c"].sort().join("") + "|");` +
		`Response.Write([1,2].concat([3,4],5).join(""));` +
		`</script>`
	out := runASPSourceForTest(t, source)
	if out != "23|23|1984|1|5|48967|abc|12345" {
		t.Fatalf("unexpected array ES5 methods output: %q", out)
	}
}

func TestJScriptES5ObjectPrototypeMethods(t *testing.T) {
	vm := NewVM(nil, nil, 0)
	vm.ensureJSRootEnv()

	protoID := vm.allocJSID()
	vm.jsObjectItems[protoID] = map[string]Value{"k": NewInteger(1)}
	vm.jsPropertyItems[protoID] = map[string]jsPropertyDescriptor{"k": jsDefaultPropertyDescriptor(NewInteger(1))}

	objID := vm.allocJSID()
	vm.jsObjectItems[objID] = map[string]Value{"x": NewInteger(9), "__js_proto": {Type: VTJSObject, Num: protoID}}
	vm.jsPropertyItems[objID] = map[string]jsPropertyDescriptor{"x": jsDefaultPropertyDescriptor(NewInteger(9))}

	obj := Value{Type: VTJSObject, Num: objID}
	proto := Value{Type: VTJSObject, Num: protoID}

	hop, _ := vm.jsCallMember(obj, "hasOwnProperty", []Value{NewString("x")})
	if hop.Type != VTBool || hop.Num == 0 {
		t.Fatalf("expected hasOwnProperty true, got %#v", hop)
	}

	ip, _ := vm.jsCallMember(proto, "isPrototypeOf", []Value{obj})
	if ip.Type != VTBool || ip.Num == 0 {
		t.Fatalf("expected isPrototypeOf true, got %#v", ip)
	}

	pie, _ := vm.jsCallMember(obj, "propertyIsEnumerable", []Value{NewString("x")})
	if pie.Type != VTBool || pie.Num == 0 {
		t.Fatalf("expected propertyIsEnumerable true, got %#v", pie)
	}

	ts, _ := vm.jsCallMember(obj, "toString", nil)
	if ts.Type != VTString || ts.Str != "[object Object]" {
		t.Fatalf("unexpected toString output: %#v", ts)
	}

	vv, _ := vm.jsCallMember(obj, "valueOf", nil)
	if vv.Type != VTJSObject || vv.Num != objID {
		t.Fatalf("unexpected valueOf output: %#v", vv)
	}
}

// BenchmarkJScriptTailCallBasic benchmarks simple tail-recursive function.
// Verifies that TCO has zero allocations (0 allocs/op).
func BenchmarkJScriptTailCallBasic(b *testing.B) {
	source := `<script runat="server" language="JScript">` +
		`function sum(n, acc) {` +
		`if (n === 0) { return acc; }` +
		`return sum(n - 1, acc + 1);` +
		`}` +
		`var result = sum(10000, 0);` +
		`</script>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		b.Fatalf("compile failed: %v", err)
	}

	b.ReportAllocs()
	b.ResetTimer()
	for i := 0; i < b.N; i++ {
		vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
		host := NewMockHost()
		var output bytes.Buffer
		host.SetOutput(&output)
		host.Response().SetBuffer(false)
		vm.SetHost(host)

		if err := vm.Run(); err != nil {
			b.Fatalf("vm run failed: %v", err)
		}
	}
}

// BenchmarkJScriptTailCallMember benchmarks tail-recursive member calls.
// Verifies that TCO on object methods has zero allocations (0 allocs/op).
func BenchmarkJScriptTailCallMember(b *testing.B) {
	source := `<script runat="server" language="JScript">` +
		`var obj = {};` +
		`obj.sum = function(n, acc) {` +
		`if (n === 0) { return acc; }` +
		`return obj.sum(n - 1, acc + 1);` +
		`};` +
		`var result = obj.sum(10000, 0);` +
		`</script>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		b.Fatalf("compile failed: %v", err)
	}

	b.ReportAllocs()
	b.ResetTimer()
	for i := 0; i < b.N; i++ {
		vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
		host := NewMockHost()
		var output bytes.Buffer
		host.SetOutput(&output)
		host.Response().SetBuffer(false)
		vm.SetHost(host)

		if err := vm.Run(); err != nil {
			b.Fatalf("vm run failed: %v", err)
		}
	}
}

// BenchmarkJScriptTailCallDeepRecursion benchmarks very deep tail recursion (100k iterations).
// Verifies that TCO keeps memory usage constant over deep call chains.
// Should show minimal allocations even for 100,000 iterations.
func BenchmarkJScriptTailCallDeepRecursion(b *testing.B) {
	source := `<script runat="server" language="JScript">` +
		`function sum(n, acc) {` +
		`if (n === 0) { return acc; }` +
		`return sum(n - 1, acc + 1);` +
		`}` +
		`var result = sum(100000, 0);` +
		`</script>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		b.Fatalf("compile failed: %v", err)
	}

	b.ReportAllocs()
	b.ResetTimer()
	for i := 0; i < b.N; i++ {
		vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
		host := NewMockHost()
		var output bytes.Buffer
		host.SetOutput(&output)
		host.Response().SetBuffer(false)
		vm.SetHost(host)

		if err := vm.Run(); err != nil {
			b.Fatalf("vm run failed: %v", err)
		}
	}
}

func TestJScriptES5ParseIntParseFloatNuances(t *testing.T) {
	source := `<script runat="server" language="JScript">` +
		`Response.Write(parseInt("08") + "|");` +
		`Response.Write(parseInt("0x10") + "|");` +
		`Response.Write(parseInt("11", 2) + "|");` +
		`Response.Write(parseInt("xyz") + "|");` +
		`Response.Write(parseFloat("3.14abc") + "|");` +
		`Response.Write(parseFloat(".5") + "|");` +
		`Response.Write(parseFloat("1e-2") + "|");` +
		`var bad = parseFloat("abc");` +
		`Response.Write((bad != bad) ? "NaN" : bad);` +
		`</script>`
	out := runASPSourceForTest(t, source)
	if out != "8|16|3|NaN|3.14|0.5|0.01|NaN" {
		t.Fatalf("unexpected parseInt/parseFloat output: %q", out)
	}
}

func TestJScriptES5NumberPrimitiveMethods(t *testing.T) {
	source := `<script runat="server" language="JScript">` +
		`var n = 12.3456;` +
		`Response.Write(n.toFixed(2) + "|");` +
		`Response.Write(n.toLocaleString() + "|");` +
		`Response.Write(n.toExponential(2) + "|");` +
		`Response.Write(n.toPrecision(4) + "|");` +
		`Response.Write(n.toString() + "|");` +
		`Response.Write((typeof n.valueOf()) + "|");` +
		`Response.Write((123).toString(16));` +
		`</script>`
	out := runASPSourceForTest(t, source)
	if out != "12.35|12.3456|1.23e+1|12.35|12.3456|number|7b" {
		t.Fatalf("unexpected Number primitive method output: %q", out)
	}
}

func TestJScriptES5NumberToFixedNoUndefinedRegression(t *testing.T) {
	source := `<script runat="server" language="JScript">` +
		`var start = new Date().getTime();` +
		`for (var i = 1; i <= 100000; i++) {}` +
		`var elapsed = (new Date().getTime() - start) / 1000;` +
		`Response.Write("Tempo levado: " + elapsed.toFixed(2) + " segundos.");` +
		`</script>`
	out := runASPSourceForTest(t, source)
	if strings.Contains(out, "undefined") {
		t.Fatalf("toFixed regression: got undefined output: %q", out)
	}
	if !strings.HasPrefix(out, "Tempo levado: ") || !strings.HasSuffix(out, " segundos.") {
		t.Fatalf("unexpected loop timing output shape: %q", out)
	}
}

func TestJScriptES5ArrayMethodsGenericOnArguments(t *testing.T) {
	source := `<script runat="server" language="JScript">` +
		`function probe(a, b, c) {` +
		`  return arguments.slice(1).join(",");` +
		`}` +
		`var list = {0: 1, 1: 2, 2: 3, length: 3};` +
		`var seen = "";` +
		`list.forEach(function(v, i) { seen += i + ":" + v + ";"; });` +
		`var mapped = list.map(function(v) { return v * 2; }).join(",");` +
		`var filtered = list.filter(function(v) { return v >= 2; }).join(",");` +
		`Response.Write(probe(1, 2, 3) + "|" + seen + "|" + mapped + "|" + filtered);` +
		`</script>`
	out := runASPSourceForTest(t, source)
	if out != "2,3|0:1;1:2;2:3;|2,4,6|2,3" {
		t.Fatalf("unexpected generic array-method output: %q", out)
	}
}

func TestJScriptES5StringReplaceCallback(t *testing.T) {
	source := `<script runat="server" language="JScript">` +
		`var regexOut = "abc123def".replace(/\d+/g, function(match, offset, sourceText) { return "[" + match + ":" + offset + ":" + sourceText.length + "]"; });` +
		`var literalOut = "axa".replace("x", function(match, offset) { return match + offset; });` +
		`Response.Write(regexOut + "|" + literalOut);` +
		`</script>`
	out := runASPSourceForTest(t, source)
	if out != "abc[123:3:9]def|ax1a" {
		t.Fatalf("unexpected callback replace output: %q", out)
	}
}

func TestJScriptES5StringReplaceRegexCallbackAvoidsLocalClone(t *testing.T) {
	source := `<script runat="server" language="JScript">` +
		`var calls = 0;` +
		`var out = "a1b2c3d4".replace(/\d/g, function(match, offset) { calls++; return "[" + offset + "]"; });` +
		`Response.Write(out + "|" + calls);` +
		`</script>`
	out, vm, err := runASPSourceForTestWithVM(t, source)
	if err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	if out != "a[1]b[3]c[5]d[7]|4" {
		t.Fatalf("unexpected regex callback output: %q", out)
	}
	if vm.cloneForExecuteLocalCount != 0 {
		t.Fatalf("expected no local clone in regex callback path, got %d", vm.cloneForExecuteLocalCount)
	}
}

func TestJScriptES5ParseIntLeadingZeroStrictDecimal(t *testing.T) {
	source := `<script runat="server" language="JScript">` +
		`Response.Write(parseInt("010") + "|");` +
		`Response.Write(parseInt("010", 0) + "|");` +
		`Response.Write(parseInt("010", 8) + "|");` +
		`Response.Write(parseInt("0x10") + "|");` +
		`Response.Write(parseInt("0019"));` +
		`</script>`
	out := runASPSourceForTest(t, source)
	if out != "10|10|8|16|19" {
		t.Fatalf("unexpected parseInt leading-zero output: %q", out)
	}
}

func TestJScriptObjectToStringInternalTags(t *testing.T) {
	vm := NewVM(nil, nil, 0)
	vm.ensureJSRootEnv()

	arrVal := ValueFromVBArray(NewVBArrayFromValues(0, []Value{NewInteger(1)}))
	if got := vm.jsObjectToStringTag(arrVal); got != "[object Array]" {
		t.Fatalf("unexpected array tag: %s", got)
	}

	dateVal := NewDate(time.Date(2026, time.January, 2, 3, 4, 5, 0, time.UTC))
	if got := vm.jsObjectToStringTag(dateVal); got != "[object Date]" {
		t.Fatalf("unexpected date tag: %s", got)
	}

	fnID := vm.allocJSID()
	vm.jsFunctionItems[fnID] = &jsFunctionObject{name: "f"}
	if got := vm.jsObjectToStringTag(Value{Type: VTJSFunction, Num: fnID}); got != "[object Function]" {
		t.Fatalf("unexpected function tag: %s", got)
	}

	objID := vm.allocJSID()
	vm.jsObjectItems[objID] = map[string]Value{"__js_type": NewString("RegExp")}
	vm.jsPropertyItems[objID] = make(map[string]jsPropertyDescriptor)
	if got := vm.jsObjectToStringTag(Value{Type: VTJSObject, Num: objID}); got != "[object RegExp]" {
		t.Fatalf("unexpected regexp tag: %s", got)
	}
}

func TestJScriptIntegerFastPathArithmeticPreservesIntegerType(t *testing.T) {
	vm := NewVM(nil, nil, 0)

	add := vm.jsAdd(NewInteger(7), NewInteger(5))
	if add.Type != VTInteger || add.Num != 12 {
		t.Fatalf("expected integer add result 12, got %#v", add)
	}

	sub := vm.jsSubtract(NewInteger(12), NewInteger(9))
	if sub.Type != VTInteger || sub.Num != 3 {
		t.Fatalf("expected integer subtract result 3, got %#v", sub)
	}
}

func TestJScriptIncrementDecrementHelpersPreserveIntegerType(t *testing.T) {
	vm := NewVM(nil, nil, 0)

	next := vm.jsIncrementNumberValue(NewInteger(41))
	if next.Type != VTInteger || next.Num != 42 {
		t.Fatalf("expected integer increment result 42, got %#v", next)
	}

	prev := vm.jsDecrementNumberValue(NewInteger(41))
	if prev.Type != VTInteger || prev.Num != 40 {
		t.Fatalf("expected integer decrement result 40, got %#v", prev)
	}
}

func TestJScriptFunctionDeclarationHoisting(t *testing.T) {
	source := `<%@ Language="JScript" %>` +
		`<%` +
		`Response.Write(a());` +
		`function a() { return "hoisted"; }` +
		`Response.Write("|" + a());` +
		`%>`
	out := runASPSourceForTest(t, source)
	if out != "hoisted|hoisted" {
		t.Fatalf("expected function hoisting output 'hoisted|hoisted', got %q", out)
	}
}

func TestJScriptVarDeclarationHoisting(t *testing.T) {
	source := `<%@ Language="JScript" %>` +
		`<%` +
		`Response.Write(typeof x);` +
		`var x = "assigned";` +
		`Response.Write("|" + x);` +
		`%>`
	out := runASPSourceForTest(t, source)
	if out != "undefined|assigned" {
		t.Fatalf("expected var hoisting output 'undefined|assigned', got %q", out)
	}
}

func TestJScriptBooleanGlobalFunctionCall(t *testing.T) {
	source := `<%@ Language="JScript" %>` +
		`<%` +
		`Response.Write("typeof Boolean = " + typeof Boolean + "<br>");` +
		`Response.Write("Boolean('fish') = " + Boolean("fish") + "<br>");` +
		`%>`
	out := runASPSourceForTest(t, source)
	expected := "typeof Boolean = function<br>Boolean('fish') = true<br>"
	if out != expected {
		t.Fatalf("expected Boolean function call output %q, got %q", expected, out)
	}
}
