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
	"os"
	"path/filepath"
	"strings"
	"testing"
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
	hasPostIncrement := false
	hasAddAssign := false
	for i := 0; i < len(bytecode); i++ {
		switch OpCode(bytecode[i]) {
		case OpJSPostIncrement:
			hasPostIncrement = true
		case OpJSAddAssign:
			hasAddAssign = true
		}
	}

	if !hasPostIncrement {
		t.Fatalf("expected OpJSPostIncrement in bytecode, got %v", bytecode)
	}
	if !hasAddAssign {
		t.Fatalf("expected OpJSAddAssign in bytecode, got %v", bytecode)
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
	source := `<script runat="server" language="JScript">` +
		`var a = 7 - 2; var b = 3 * 4; var c = 9 / 3; var d = 9 % 4;` +
		`var e = ("5" == 5); var f = (2 < "10");` +
		`Response.Write(a + b + c + d + (e ? 1 : 0) + (f ? 1 : 0));` +
		`</script>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	bytecode := compiler.Bytecode()
	hasJSSubtract := false
	hasJSMultiply := false
	hasJSDivide := false
	hasJSModulo := false
	hasJSLooseEq := false
	hasJSLess := false
	for i := 0; i < len(bytecode); i++ {
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
		t.Fatalf("expected JS binary opcodes in bytecode, got %v", bytecode)
	}

	out := runASPSourceForTest(t, source)
	if out != "23" {
		t.Fatalf("unexpected binary operator output: %q", out)
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
	if out != "strict=False|and=False|count=0" {
		t.Fatalf("unexpected null/false/zero rendering output: %q", out)
	}
}
