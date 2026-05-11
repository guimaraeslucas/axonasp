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
	"testing"
)

func benchmarkASPExecutionOnly(b *testing.B, source string) {
	b.Helper()
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

func TestJScriptLoopOnlyCases(t *testing.T) {
	testCases := []struct {
		name   string
		source string
		want   string
	}{
		{
			name: "while loop",
			source: `<script runat="server" language="JScript">` +
				`var i = 0;` +
				`var out = "";` +
				`while (i < 3) { out += i + ","; i++; }` +
				`Response.Write(out);` +
				`</script>`,
			want: "0,1,2,",
		},
		{
			name: "do while loop",
			source: `<script runat="server" language="JScript">` +
				`var i = 0;` +
				`var out = "";` +
				`do { out += i + ","; i++; } while (i < 2);` +
				`Response.Write(out);` +
				`</script>`,
			want: "0,1,",
		},
		{
			name: "for loop",
			source: `<script runat="server" language="JScript">` +
				`var out = "";` +
				`for (var i = 0; i < 3; i++) { out += i + ","; }` +
				`Response.Write(out);` +
				`</script>`,
			want: "0,1,2,",
		},
		{
			name: "for break",
			source: `<script runat="server" language="JScript">` +
				`var out = "";` +
				`for (var i = 0; i < 5; i++) { if (i === 3) { break; } out += i + ","; }` +
				`Response.Write(out);` +
				`</script>`,
			want: "0,1,2,",
		},
		{
			name: "for continue",
			source: `<script runat="server" language="JScript">` +
				`var out = "";` +
				`for (var i = 0; i < 5; i++) { if (i === 2) { continue; } out += i + ","; }` +
				`Response.Write(out);` +
				`</script>`,
			want: "0,1,3,4,",
		},
		{
			name: "switch inside loop",
			source: `<script runat="server" language="JScript">` +
				`var out = "";` +
				`for (var i = 0; i < 4; i++) {` +
				`  switch (i) {` +
				`  case 0: out += "a"; break;` +
				`  case 1: continue;` +
				`  case 2: out += "c"; break;` +
				`  default: out += "d";` +
				`  }` +
				`}` +
				`Response.Write(out);` +
				`</script>`,
			want: "acd",
		},
		{
			name: "for in loop",
			source: `<script runat="server" language="JScript">` +
				`var o = {};` +
				`o.b = 2; o.a = 1; o.c = 3;` +
				`var out = "";` +
				`for (var key in o) { if (key === "b") { continue; } out += key; if (key === "c") { break; } }` +
				`Response.Write(out);` +
				`</script>`,
			want: "ac",
		},
		{
			name: "percent block for loop",
			source: `<%@ Language="JScript" %><%` +
				`var out = "";` +
				`for (var i = 0; i < 3; i++) { out += i; }` +
				`Response.Write(out);` +
				`%>`,
			want: "012",
		},
	}

	for _, tc := range testCases {
		t.Run(tc.name, func(t *testing.T) {
			out := runASPSourceForTest(t, tc.source)
			if out != tc.want {
				t.Fatalf("unexpected output: got %q want %q", out, tc.want)
			}
		})
	}
}

func TestJScriptForInLoopStateResetsAcrossFunctionCalls(t *testing.T) {
	source := `<script runat="server" language="JScript">` +
		`function iterateOnce(obj) {` +
		`  var out = "";` +
		`  for (var key in obj) { out += key; break; }` +
		`  return out;` +
		`}` +
		`var o = {};` +
		`o.a = 1; o.b = 2; o.c = 3;` +
		`Response.Write(iterateOnce(o) + "|" + iterateOnce(o));` +
		`</script>`

	out := runASPSourceForTest(t, source)
	if out != "a|a" {
		t.Fatalf("unexpected repeated for-in output: %q", out)
	}
}

func TestJScriptForInLoopInsideFunctionSingleCall(t *testing.T) {
	source := `<script runat="server" language="JScript">` +
		`function iterateOnce(obj) {` +
		`  var out = "";` +
		`  for (var key in obj) { out += key; break; }` +
		`  return out;` +
		`}` +
		`var o = {};` +
		`o.a = 1; o.b = 2; o.c = 3;` +
		`Response.Write(iterateOnce(o));` +
		`</script>`

	out := runASPSourceForTest(t, source)
	if out != "a" {
		t.Fatalf("unexpected single-call for-in output: %q", out)
	}
}

func TestJScriptForInLoopInsideFunctionWithoutBreak(t *testing.T) {
	source := `<script runat="server" language="JScript">` +
		`function iterateAll(obj) {` +
		`  var out = "";` +
		`  for (var key in obj) { out += key; }` +
		`  return out;` +
		`}` +
		`var o = {};` +
		`o.a = 1; o.b = 2; o.c = 3;` +
		`Response.Write(iterateAll(o));` +
		`</script>`

	out := runASPSourceForTest(t, source)
	if out != "abc" {
		t.Fatalf("unexpected full for-in output: %q", out)
	}
}

func TestJScriptLoopStressTerminates(t *testing.T) {
	source := `<script runat="server" language="JScript">` +
		`var i = 0;` +
		`while (i < 5000) { i++; }` +
		`Response.Write(i);` +
		`</script>`

	out := runASPSourceForTest(t, source)
	if out != "5000" {
		t.Fatalf("unexpected stress-loop output: %q", out)
	}
}

func TestJScriptObjectArgumentAccessInsideFunction(t *testing.T) {
	source := `<script runat="server" language="JScript">` +
		`function inspect(obj) { return typeof obj + "|" + obj.a + "|" + obj.b; }` +
		`var o = {};` +
		`o.a = 1; o.b = 2;` +
		`Response.Write(inspect(o));` +
		`</script>`

	out := runASPSourceForTest(t, source)
	if out != "object|1|2" {
		t.Fatalf("unexpected object-argument output: %q", out)
	}
}

func TestJScriptForInLoopInsideFunctionCanReturnFromBody(t *testing.T) {
	source := `<script runat="server" language="JScript">` +
		`function firstKey(obj) {` +
		`  for (var key in obj) { return "hit:" + key; }` +
		`  return "none";` +
		`}` +
		`var o = {};` +
		`o.a = 1; o.b = 2;` +
		`Response.Write(firstKey(o));` +
		`</script>`

	out := runASPSourceForTest(t, source)
	if out != "hit:a" {
		t.Fatalf("unexpected for-in body return output: %q", out)
	}
}

func TestJScriptTopLevelForLetUsesFastIntOpcode(t *testing.T) {
	source := `<%@ Language="JScript" %><%` +
		`var sum = 0;` +
		`for (let i = 0; i < 100; i++) { sum = sum + i; }` +
		`Response.Write(sum);` +
		`%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	hasFastInt := false
	hasIterEnter := false
	for i := 0; i < len(compiler.Bytecode()); i++ {
		switch OpCode(compiler.Bytecode()[i]) {
		case OpJSForFastInt:
			hasFastInt = true
		case OpJSForIterEnter:
			hasIterEnter = true
		}
	}
	if !hasFastInt {
		t.Fatalf("expected OpJSForFastInt in bytecode, got %v", compiler.Bytecode())
	}
	if hasIterEnter {
		t.Fatalf("did not expect OpJSForIterEnter for top-level fast-int loop, got %v", compiler.Bytecode())
	}

	out := runASPSourceForTest(t, source)
	if out != "4950" {
		t.Fatalf("unexpected output: %q", out)
	}
}

func TestJScriptTopLevelFastIntUsesRootFrameOpcodes(t *testing.T) {
	source := `<%@ Language="JScript" %><%` +
		`var sum = 0;` +
		`for (let i = 0; i < 8; i++) { sum = sum + i; }` +
		`Response.Write(sum);` +
		`%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	hasRootEnter := false
	hasRootLeave := false
	for i := 0; i < len(compiler.Bytecode()); i++ {
		switch OpCode(compiler.Bytecode()[i]) {
		case OpJSRootFrameEnter:
			hasRootEnter = true
		case OpJSRootFrameLeave:
			hasRootLeave = true
		}
	}
	if !hasRootEnter || !hasRootLeave {
		t.Fatalf("expected root frame opcodes in bytecode, got %v", compiler.Bytecode())
	}

	out := runASPSourceForTest(t, source)
	if out != "28" {
		t.Fatalf("unexpected output: %q", out)
	}
}

func TestJScriptTopLevelVarRemainsGlobalWithLoopLet(t *testing.T) {
	source := `<%@ Language="JScript" %><%` +
		`var i = 99;` +
		`for (let i = 0; i < 3; i++) {}` +
		`Response.Write(i);` +
		`%>`

	out := runASPSourceForTest(t, source)
	if out != "99" {
		t.Fatalf("unexpected output: %q", out)
	}
}

func TestJScriptForLetNoCaptureUsesFastIterOpcodes(t *testing.T) {
	source := `<%@ Language="JScript" %><%` +
		`var sum = 0;` +
		`for (let i = 0; i < 10; i = i + 1) { sum = sum + i; }` +
		`Response.Write(sum);` +
		`%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	hasFastIterEnter := false
	hasFastIterExit := false
	hasSlowIterEnter := false
	for i := 0; i < len(compiler.Bytecode()); i++ {
		switch OpCode(compiler.Bytecode()[i]) {
		case OpJSForIterEnterFast:
			hasFastIterEnter = true
		case OpJSForIterExitFast:
			hasFastIterExit = true
		case OpJSForIterEnter:
			hasSlowIterEnter = true
		}
	}
	if !hasFastIterEnter || !hasFastIterExit {
		t.Fatalf("expected fast iter opcodes in bytecode, got %v", compiler.Bytecode())
	}
	if hasSlowIterEnter {
		t.Fatalf("did not expect slow iter enter opcode in non-capturing loop, got %v", compiler.Bytecode())
	}

	out := runASPSourceForTest(t, source)
	if out != "45" {
		t.Fatalf("unexpected output: %q", out)
	}
}

func TestJScriptForLetCaptureUsesSlowIterOpcode(t *testing.T) {
	source := `<%@ Language="JScript" %><%` +
		`var funcs = [];` +
		`for (let i = 0; i < 4; i = i + 1) { funcs.push(function(){ return i; }); }` +
		`Response.Write(funcs[0]() + "," + funcs[1]() + "," + funcs[2]() + "," + funcs[3]());` +
		`%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	hasSlowIterEnter := false
	for i := 0; i < len(compiler.Bytecode()); i++ {
		if OpCode(compiler.Bytecode()[i]) == OpJSForIterEnter {
			hasSlowIterEnter = true
			break
		}
	}
	if !hasSlowIterEnter {
		t.Fatalf("expected slow iter opcode for capturing loop, got %v", compiler.Bytecode())
	}
}

func BenchmarkJScriptTopLevelFastInt1M(b *testing.B) {
	source := `<%@ Language="JScript" %><%` +
		`var sum = 0;` +
		`for (let i = 0; i < 1000000; i++) { sum = sum + i; }` +
		`Response.Write(sum);` +
		`%>`
	benchmarkASPExecutionOnly(b, source)
}

func BenchmarkJScriptTopLevelFallbackNoCapture1M(b *testing.B) {
	source := `<%@ Language="JScript" %><%` +
		`var sum = 0;` +
		`for (let i = 0; i < 1000000; i = i + 1) { sum = sum + i; }` +
		`Response.Write(sum);` +
		`%>`
	benchmarkASPExecutionOnly(b, source)
}

func BenchmarkJScriptTopLevelFallbackCapture100K(b *testing.B) {
	source := `<%@ Language="JScript" %><%` +
		`var funcs = [];` +
		`for (let i = 0; i < 100000; i = i + 1) { funcs.push(function(){ return i; }); }` +
		`Response.Write(funcs[0]());` +
		`%>`
	benchmarkASPExecutionOnly(b, source)
}
