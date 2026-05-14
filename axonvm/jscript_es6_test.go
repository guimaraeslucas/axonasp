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
	"bytes"
	"os"
	"path/filepath"
	"strings"
	"testing"
)

func runJScriptModuleEntry(t *testing.T, entryPath string) (string, error) {
	t.Helper()

	cache := getExecuteScriptCache()
	program, err := cache.LoadOrCompile(entryPath)
	if err != nil {
		return "", err
	}

	vm := NewVMFromCachedProgram(program)
	vm.sourceName = entryPath
	vm.baseSourceName = entryPath
	host := NewMockHost()
	var out bytes.Buffer
	host.SetOutput(&out)
	host.Response().SetBuffer(false)
	vm.SetHost(host)

	err = vm.Run()
	return out.String(), err
}

func TestJScriptModuleCache(t *testing.T) {
	// 1. Create a dummy .js module file
	jsPath := filepath.Join(t.TempDir(), "test_module.js")
	jsCode := "var x = 42; Response.Write(x);"
	if err := os.WriteFile(jsPath, []byte(jsCode), 0644); err != nil {
		t.Fatal(err)
	}

	// 2. Get the global execute cache (which we modified to use NewJSModuleCompiler)
	cache := getExecuteScriptCache()

	// 3. Load or compile it
	program, err := cache.LoadOrCompile(jsPath)
	if err != nil {
		t.Fatal(err)
	}

	// 4. Verify it's a "module" by looking at the source name or just running it
	if program.SourceName != normalizeScriptCacheKey(jsPath) {
		t.Errorf("expected source name %q, got %q", jsPath, program.SourceName)
	}

	// 5. Run it
	vm := NewVMFromCachedProgram(program)
	host := NewMockHost()
	var out bytes.Buffer
	host.SetOutput(&out)
	host.Response().SetBuffer(false)
	vm.SetHost(host)
	if err := vm.Run(); err != nil {
		t.Fatal(err)
	}
	if out.String() != "42" {
		t.Errorf("expected '42', got %q", out.String())
	}
}

func TestJScriptModuleImportExportNamed(t *testing.T) {
	dir := t.TempDir()
	depPath := filepath.Join(dir, "dep.js")
	entryPath := filepath.Join(dir, "entry.js")

	depSrc := `export var value = 7; export function inc(x) { return x + 1; }`
	entrySrc := `import { value, inc } from "./dep.js"; Response.Write(inc(value));`

	if err := os.WriteFile(depPath, []byte(depSrc), 0644); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(entryPath, []byte(entrySrc), 0644); err != nil {
		t.Fatal(err)
	}

	out, err := runJScriptModuleEntry(t, entryPath)
	if err != nil {
		t.Fatal(err)
	}
	if out != "8" {
		t.Fatalf("expected 8, got %q", out)
	}
}

func TestJScriptModuleReExportFrom(t *testing.T) {
	dir := t.TempDir()
	aPath := filepath.Join(dir, "a.js")
	bPath := filepath.Join(dir, "b.js")
	entryPath := filepath.Join(dir, "entry.js")

	if err := os.WriteFile(aPath, []byte(`export var n = 10;`), 0644); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(bPath, []byte(`export { n as count } from "./a.js";`), 0644); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(entryPath, []byte(`import { count } from "./b.js"; Response.Write(count);`), 0644); err != nil {
		t.Fatal(err)
	}

	out, err := runJScriptModuleEntry(t, entryPath)
	if err != nil {
		t.Fatal(err)
	}
	if out != "10" {
		t.Fatalf("expected 10, got %q", out)
	}
}

func TestJScriptModuleCircularImport(t *testing.T) {
	dir := t.TempDir()
	aPath := filepath.Join(dir, "a.js")
	bPath := filepath.Join(dir, "b.js")
	entryPath := filepath.Join(dir, "entry.js")

	aSrc := `import { b } from "./b.js"; export var a = "A"; export var fromB = b;`
	bSrc := `import { a } from "./a.js"; export var b = "B"; export var fromA = a;`
	entrySrc := `import { a, fromB } from "./a.js"; import { b, fromA } from "./b.js"; Response.Write(a + b + "|" + fromA + "|" + fromB);`

	if err := os.WriteFile(aPath, []byte(aSrc), 0644); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(bPath, []byte(bSrc), 0644); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(entryPath, []byte(entrySrc), 0644); err != nil {
		t.Fatal(err)
	}

	out, err := runJScriptModuleEntry(t, entryPath)
	if err != nil {
		t.Fatal(err)
	}
	if out != "AB|undefined|B" {
		t.Fatalf("expected AB|undefined|B, got %q", out)
	}
}

func TestJScriptModuleCanUseASPObjects(t *testing.T) {
	dir := t.TempDir()
	depPath := filepath.Join(dir, "dep.js")
	entryPath := filepath.Join(dir, "entry.js")

	if err := os.WriteFile(depPath, []byte(`Response.Write("M"); export var ok = 1;`), 0644); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(entryPath, []byte(`import "./dep.js"; import { ok } from "./dep.js"; Response.Write(ok);`), 0644); err != nil {
		t.Fatal(err)
	}

	out, err := runJScriptModuleEntry(t, entryPath)
	if err != nil {
		t.Fatal(err)
	}
	if out != "M1" {
		t.Fatalf("expected M1, got %q", out)
	}
}

func TestJScriptModuleOpcodeEmission(t *testing.T) {
	compiler := NewJSModuleCompiler(`import { value } from "./dep.js"; export var x = value;`)
	if err := compiler.Compile(); err != nil {
		t.Fatal(err)
	}

	hasImport := false
	hasExport := false
	for ip := 0; ip < len(compiler.bytecode); {
		op := OpCode(compiler.bytecode[ip])
		ip++
		if op == OpJSImport {
			hasImport = true
		}
		if op == OpJSExport {
			hasExport = true
		}
		size := opcodeOperandSize(op)
		if op == OpJSImport {
			if ip+4 > len(compiler.bytecode) {
				t.Fatal("invalid JSImport bytecode")
			}
			specCount := int(compiler.bytecode[ip+2])<<8 | int(compiler.bytecode[ip+3])
			size = 4 + (specCount * 4)
		}
		if op == OpJSForIterEnter || op == OpJSForIterExit {
			if ip+2 > len(compiler.bytecode) {
				t.Fatal("invalid for-iter bytecode")
			}
			count := int(compiler.bytecode[ip])<<8 | int(compiler.bytecode[ip+1])
			size = 2 + count*2
		}
		ip += size
	}

	if !hasImport {
		t.Fatal("expected OpJSImport in module bytecode")
	}
	if !hasExport {
		t.Fatal("expected OpJSExport in module bytecode")
	}
}

func TestJScriptModuleExportOnlyRuns(t *testing.T) {
	dir := t.TempDir()
	entryPath := filepath.Join(dir, "export_only.js")
	src := `export var value = 7; export function inc(x) { return x + 1; }`
	if err := os.WriteFile(entryPath, []byte(src), 0644); err != nil {
		t.Fatal(err)
	}

	_, err := runJScriptModuleEntry(t, entryPath)
	if err != nil {
		t.Fatal(err)
	}
}

func TestJScriptModuleImportSideEffectFromExportModule(t *testing.T) {
	dir := t.TempDir()
	depPath := filepath.Join(dir, "dep.js")
	entryPath := filepath.Join(dir, "entry.js")

	if err := os.WriteFile(depPath, []byte(`export var value = 7; Response.Write("D");`), 0644); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(entryPath, []byte(`import "./dep.js"; Response.Write("E");`), 0644); err != nil {
		t.Fatal(err)
	}

	out, err := runJScriptModuleEntry(t, entryPath)
	if err != nil {
		t.Fatal(err)
	}
	if out != "DE" {
		t.Fatalf("expected DE, got %q", out)
	}
}

// ---------------------------------------------------------------------------
// ES6 Template Literals
// ---------------------------------------------------------------------------

func TestJScriptTemplateLiteralPlain(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		var s = `+"`"+`hello world`+"`"+`;
		Response.Write(s);
	`))
	if err != nil {
		t.Fatal(err)
	}
	if out != "hello world" {
		t.Errorf("expected 'hello world', got %q", out)
	}
}

func TestJScriptTemplateLiteralInterpolation(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		var name = "World";
		var s = `+"`"+`Hello ${name}!`+"`"+`;
		Response.Write(s);
	`))
	if err != nil {
		t.Fatal(err)
	}
	if out != "Hello World!" {
		t.Errorf("expected 'Hello World!', got %q", out)
	}
}

func TestJScriptTemplateLiteralExpression(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		var a = 3;
		var b = 4;
		var s = `+"`"+`${a} + ${b} = ${a + b}`+"`"+`;
		Response.Write(s);
	`))
	if err != nil {
		t.Fatal(err)
	}
	if out != "3 + 4 = 7" {
		t.Errorf("expected '3 + 4 = 7', got %q", out)
	}
}

func TestJScriptTemplateLiteralMultiline(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		var s = `+"`"+`line1
line2`+"`"+`;
		Response.Write(s);
	`))
	if err != nil {
		t.Fatal(err)
	}
	if !strings.Contains(out, "line1") || !strings.Contains(out, "line2") {
		t.Errorf("expected multiline output, got %q", out)
	}
}

func TestJScriptGenerators(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		function* gen() {
			yield 1;
			yield 2;
			return 3;
		}
		var g = gen();
		var r1 = g.next();
		var r2 = g.next();
		var r3 = g.next();
		Response.Write(r1.value + ":" + r1.done + "|");
		Response.Write(r2.value + ":" + r2.done + "|");
		Response.Write(r3.value + ":" + r3.done);
	`))
	if err != nil {
		t.Fatal(err)
	}
	expected := `1:False|2:False|3:True`
	if out != expected {
		t.Errorf("expected %q, got %q", expected, out)
	}
}

func TestJScriptGeneratorsResume(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		function* gen() {
			var x = yield 1;
			var y = yield (x + 10);
			return x + y;
		}
		var g = gen();
		var r1 = g.next();
		var r2 = g.next(5);
		var r3 = g.next(10);
		Response.Write(r1.value + "|" + r2.value + "|" + r3.value);
	`))
	if err != nil {
		t.Fatal(err)
	}
	expected := `1|15|15`
	if out != expected {
		t.Errorf("expected %q, got %q", expected, out)
	}
}

func TestJScriptGeneratorLoopConstant(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		function* idMaker() {
			while (true)
				yield 1;
		}

		var gen = idMaker();
		Response.Write(gen.next().value + "|");
		Response.Write(gen.next().value + "|");
		Response.Write(gen.next().value);
	`))
	if err != nil {
		t.Fatal(err)
	}
	if out != "1|1|1" {
		t.Errorf("expected '1|1|1', got %q", out)
	}
}

func TestJScriptGeneratorLoopPostIncrement(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		function* idMaker() {
			var index = 0;
			while (true)
				yield index++;
		}

		var gen = idMaker();
		Response.Write(gen.next().value + "|");
		Response.Write(gen.next().value + "|");
		Response.Write(gen.next().value);
	`))
	if err != nil {
		t.Fatal(err)
	}
	if out != "0|1|2" {
		t.Errorf("expected '0|1|2', got %q", out)
	}
}

func TestJScriptAsyncAwait(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		async function f() {
			var p = Promise.resolve(42);
			var val = await p;
			return val + 1;
		}
		f().then(function(v) {
			Response.Write(v);
		});
	`))
	if err != nil {
		t.Fatal(err)
	}
	if out != "43" {
		t.Errorf("expected '43', got %q", out)
	}
}

func TestJScriptAsyncAwaitReject(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		async function f() {
			await Promise.reject("fail");
			return "ok";
		}
		f().catch(function(e) {
			Response.Write("Caught: " + e);
		});
	`))
	if err != nil {
		t.Fatal(err)
	}
	if !strings.Contains(out, "Caught: fail") {
		t.Errorf("expected output to contain 'Caught: fail', got %q", out)
	}
}

func TestJScriptTemplateLiteralEscapeSequences(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		var s = `+"`"+`A\nB\tC`+"`"+`;
		var ok = s.length === 5 && s.charCodeAt(1) === 10 && s.charCodeAt(3) === 9;
		Response.Write(ok);
	`))
	if err != nil {
		t.Fatal(err)
	}
	if out != "True" {
		t.Errorf("expected template literal escapes to decode, got %q", out)
	}
}

// ---------------------------------------------------------------------------
// ES6 Arrow Functions
// ---------------------------------------------------------------------------

func TestJScriptArrowFunctionConcise(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		var double = x => x * 2;
		Response.Write(double(5));
	`))
	if err != nil {
		t.Fatal(err)
	}
	if out != "10" {
		t.Errorf("expected '10', got %q", out)
	}
}

func TestJScriptArrowFunctionBlock(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		var add = (a, b) => { return a + b; };
		Response.Write(add(3, 4));
	`))
	if err != nil {
		t.Fatal(err)
	}
	if out != "7" {
		t.Errorf("expected '7', got %q", out)
	}
}

func TestJScriptArrowFunctionLexicalThis(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		function Counter() {
			this.count = 0;
			this.inc = function() {
				var fn = () => { this.count = this.count + 1; };
				fn();
			};
		}
		var c = new Counter();
		c.inc();
		c.inc();
		Response.Write(c.count);
	`))
	if err != nil {
		t.Fatal(err)
	}
	if out != "2" {
		t.Errorf("expected '2', got %q", out)
	}
}

// ---------------------------------------------------------------------------
// ES6 Default Parameter Values
// ---------------------------------------------------------------------------

func TestJScriptDefaultParams(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		function greet(name, greeting) {
			if (greeting === undefined) greeting = "Hello";
			return greeting + ", " + name + "!";
		}
		Response.Write(greet("World"));
		Response.Write("|");
		Response.Write(greet("Alice", "Hi"));
	`))
	if err != nil {
		t.Fatal(err)
	}
	if out != "Hello, World!|Hi, Alice!" {
		t.Errorf("expected 'Hello, World!|Hi, Alice!', got %q", out)
	}
}

func TestJScriptDefaultParamSyntax(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		function multiply(a, b = 2) {
			return a * b;
		}
		Response.Write(multiply(5));
		Response.Write("|");
		Response.Write(multiply(5, 3));
	`))
	if err != nil {
		t.Fatal(err)
	}
	if out != "10|15" {
		t.Errorf("expected '10|15', got %q", out)
	}
}

// ---------------------------------------------------------------------------
// ES6 String Methods
// ---------------------------------------------------------------------------

func TestJScriptStringIncludes(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		var s = "Hello World";
		Response.Write(s.includes("World") ? "yes" : "no");
		Response.Write("|");
		Response.Write(s.includes("xyz") ? "yes" : "no");
	`))
	if err != nil {
		t.Fatal(err)
	}
	if out != "yes|no" {
		t.Errorf("expected 'yes|no', got %q", out)
	}
}

func TestJScriptStringStartsWith(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		var s = "Hello World";
		Response.Write(s.startsWith("Hello") ? "yes" : "no");
		Response.Write("|");
		Response.Write(s.startsWith("World") ? "yes" : "no");
	`))
	if err != nil {
		t.Fatal(err)
	}
	if out != "yes|no" {
		t.Errorf("expected 'yes|no', got %q", out)
	}
}

func TestJScriptStringEndsWith(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		var s = "Hello World";
		Response.Write(s.endsWith("World") ? "yes" : "no");
		Response.Write("|");
		Response.Write(s.endsWith("Hello") ? "yes" : "no");
	`))
	if err != nil {
		t.Fatal(err)
	}
	if out != "yes|no" {
		t.Errorf("expected 'yes|no', got %q", out)
	}
}

func TestJScriptStringRepeat(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		var s = "ab";
		Response.Write(s.repeat(3));
	`))
	if err != nil {
		t.Fatal(err)
	}
	if out != "ababab" {
		t.Errorf("expected 'ababab', got %q", out)
	}
}

func TestJScriptStringCodePointAt(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		var s = "A😀B";
		Response.Write(s.codePointAt(0));
		Response.Write("|");
		Response.Write(s.codePointAt(1));
		Response.Write("|");
		Response.Write(s.codePointAt(2));
		Response.Write("|");
		Response.Write(s.codePointAt(99) === undefined ? "undef" : "bad");
	`))
	if err != nil {
		t.Fatal(err)
	}
	if out != "65|128512|56832|undef" {
		t.Errorf("expected '65|128512|56832|undef', got %q", out)
	}
}

func TestJScriptStringNormalize(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		var decomposed = "e\u0301";
		Response.Write(decomposed.normalize("NFC") === "é" ? "yes" : "no");
		Response.Write("|");
		Response.Write("é".normalize("NFD") === decomposed ? "yes" : "no");
		Response.Write("|");
		try {
			"x".normalize("BAD");
			Response.Write("noerr");
		} catch (e) {
			Response.Write(("" + e).indexOf("RangeError") >= 0 ? "range" : "other");
		}
	`))
	if err != nil {
		t.Fatal(err)
	}
	if out != "yes|yes|range" {
		t.Errorf("expected 'yes|yes|range', got %q", out)
	}
}

// ---------------------------------------------------------------------------
// ES6 Number Static Methods
// ---------------------------------------------------------------------------

func TestJScriptNumberIsInteger(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		Response.Write(Number.isInteger(42) ? "yes" : "no");
		Response.Write("|");
		Response.Write(Number.isInteger(42.5) ? "yes" : "no");
		Response.Write("|");
		Response.Write(Number.isInteger("42") ? "yes" : "no");
	`))
	if err != nil {
		t.Fatal(err)
	}
	if out != "yes|no|no" {
		t.Errorf("expected 'yes|no|no', got %q", out)
	}
}

func TestJScriptNumberIsNaN(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		Response.Write(Number.isNaN(NaN) ? "yes" : "no");
		Response.Write("|");
		Response.Write(Number.isNaN(42) ? "yes" : "no");
		Response.Write("|");
		Response.Write(Number.isNaN("NaN") ? "yes" : "no");
	`))
	if err != nil {
		t.Fatal(err)
	}
	if out != "yes|no|no" {
		t.Errorf("expected 'yes|no|no', got %q", out)
	}
}

func TestJScriptNumberIsFinite(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		Response.Write(Number.isFinite(42) ? "yes" : "no");
		Response.Write("|");
		Response.Write(Number.isFinite(Infinity) ? "yes" : "no");
		Response.Write("|");
		Response.Write(Number.isFinite("42") ? "yes" : "no");
	`))
	if err != nil {
		t.Fatal(err)
	}
	if out != "yes|no|no" {
		t.Errorf("expected 'yes|no|no', got %q", out)
	}
}

func TestJScriptNumberIsSafeInteger(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		Response.Write(Number.isSafeInteger(9007199254740991) ? "yes" : "no");
		Response.Write("|");
		Response.Write(Number.isSafeInteger(9007199254740992) ? "yes" : "no");
		Response.Write("|");
		Response.Write(Number.isSafeInteger(42.5) ? "yes" : "no");
	`))
	if err != nil {
		t.Fatal(err)
	}
	if out != "yes|no|no" {
		t.Errorf("expected 'yes|no|no', got %q", out)
	}
}

func TestJScriptObjectStaticsPhase3(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		var s = Symbol("secret");
		var proto = { x: 7 };
		var obj = {};
		Object.setPrototypeOf(obj, proto);
		obj[s] = 42;

		var syms = Object.getOwnPropertySymbols(obj);
		var sameProto = Object.getPrototypeOf(obj) === proto;
		var objectIs = Object.is(NaN, NaN) && !Object.is(0, -0) && Object.is(-0, -0);

		Response.Write(objectIs ? "yes" : "no");
		Response.Write("|");
		Response.Write(sameProto ? "yes" : "no");
		Response.Write("|");
		Response.Write(syms.length);
		Response.Write("|");
		Response.Write(syms[0] === s ? "yes" : "no");
		Response.Write("|");
		Response.Write(obj.x);
	`))
	if err != nil {
		t.Fatal(err)
	}
	if out != "yes|yes|1|yes|7" {
		t.Errorf("expected 'yes|yes|1|yes|7', got %q", out)
	}

	_, err = runJScript2(t, jscriptSrc(`
		var frozen = {};
		Object.preventExtensions(frozen);
		Object.setPrototypeOf(frozen, { y: 1 });
	`))
	if err == nil {
		t.Fatal("expected TypeError when changing prototype of non-extensible object")
	}
}

// ---------------------------------------------------------------------------
// ES6 Secondary Features
// ---------------------------------------------------------------------------

func TestJScriptCompoundAssignmentOperators2(t *testing.T) {
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

func TestJScriptArraySpreadLiteral(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		var other = [3, 4];
		var arr = [1, 2, ...other, 5];
		Response.Write(arr.join(","));
	`))
	if err != nil {
		t.Fatal(err)
	}
	if out != "1,2,3,4,5" {
		t.Errorf("expected '1,2,3,4,5', got %q", out)
	}
}

func TestJScriptArrayFindAndFindIndex(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		var arr = [3, 7, 11, 14];
		var v = arr.find(function(x) { return x > 10; });
		var i = arr.findIndex(function(x) { return x > 10; });
		var missV = arr.find(function(x) { return x > 99; });
		var missI = arr.findIndex(function(x) { return x > 99; });
		Response.Write(v + "|" + i + "|" + (missV === undefined ? "undef" : missV) + "|" + missI);
	`))
	if err != nil {
		t.Fatal(err)
	}
	if out != "11|2|undef|-1" {
		t.Errorf("expected '11|2|undef|-1', got %q", out)
	}
}

func TestJScriptArrayKeysEntriesIterators(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		var arr = [10, 20, 30];
		var keys = [];
		for (var k of arr.keys()) {
			keys.push(k);
		}
		var entries = [];
		for (var e of arr.entries()) {
			entries.push(e[0] + ":" + e[1]);
		}
		Response.Write(keys.join(","));
		Response.Write("|");
		Response.Write(entries.join(","));
	`))
	if err != nil {
		t.Fatal(err)
	}
	if out != "0,1,2|0:10,1:20,2:30" {
		t.Errorf("expected '0,1,2|0:10,1:20,2:30', got %q", out)
	}
}

func TestJScriptArrayCopyWithinPhase2(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		var a = [1, 2, 3, 4, 5];
		a.copyWithin(0, 3);
		var b = [1, 2, 3, 4, 5];
		b.copyWithin(-2, 0, 2);
		Response.Write(a.join(","));
		Response.Write("|");
		Response.Write(b.join(","));
	`))
	if err != nil {
		t.Fatal(err)
	}
	if out != "4,5,3,4,5|1,2,3,1,2" {
		t.Errorf("expected '4,5,3,4,5|1,2,3,1,2', got %q", out)
	}
}

func TestJScriptBinaryAndOctalLiterals(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		var b = 0b1010;
		var o = 0o744;
		Response.Write(b + "|" + o);
	`))
	if err != nil {
		t.Fatal(err)
	}
	if out != "10|484" {
		t.Errorf("expected '10|484', got %q", out)
	}
}

func TestJScriptMathTruncSignCbrt(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		Response.Write(Math.trunc(4.9));
		Response.Write("|");
		Response.Write(Math.sign(-42));
		Response.Write("|");
		Response.Write(Math.sign(0));
		Response.Write("|");
		Response.Write(Math.cbrt(27));
	`))
	if err != nil {
		t.Fatal(err)
	}
	if out != "4|-1|0|3" {
		t.Errorf("expected '4|-1|0|3', got %q", out)
	}
}

func TestJScriptGlobalURIFunctions(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		var full = "https://example.com/a path/?q=hello world&x=1+2#frag";
		var encoded = encodeURI(full);
		Response.Write(encoded === "https://example.com/a%20path/?q=hello%20world&x=1+2#frag" ? "ok" : "bad");
		Response.Write("|");
		Response.Write(decodeURI(encoded) === full ? "ok" : "bad");
		Response.Write("|");

		var component = "q=hello world&x=1+2";
		var encodedComponent = encodeURIComponent(component);
		Response.Write(encodedComponent === "q%3Dhello%20world%26x%3D1%2B2" ? "ok" : "bad");
		Response.Write("|");
		Response.Write(decodeURIComponent(encodedComponent) === component ? "ok" : "bad");
		Response.Write("|");

		try {
			decodeURIComponent("%");
			Response.Write("noerror");
		} catch (e) {
			Response.Write(e.name || "error");
		}
	`))
	if err != nil {
		t.Fatal(err)
	}
	if out != "ok|ok|ok|ok|TypeError" {
		t.Errorf("expected 'ok|ok|ok|ok|TypeError', got %q", out)
	}
}

func TestJScriptMathPhase1Extensions(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		Response.Write(Math.acosh(1));
		Response.Write("|");
		Response.Write(Math.asinh(0));
		Response.Write("|");
		Response.Write(Math.atanh(0.5).toFixed(3));
		Response.Write("|");
		Response.Write(Math.expm1(1).toFixed(6));
		Response.Write("|");
		Response.Write(Math.log1p(1).toFixed(6));
		Response.Write("|");
		Response.Write(Math.log10(1000));
		Response.Write("|");
		Response.Write(Math.log2(8));
		Response.Write("|");
		Response.Write(Math.hypot(3, 4));
		Response.Write("|");
		Response.Write(Math.imul(0xffffffff, 5));
		Response.Write("|");
		Response.Write(Math.clz32(1));
		Response.Write("|");
		Response.Write(Math.clz32(0));
		Response.Write("|");
		Response.Write(isNaN(Math.log1p(-2)) ? "nan" : "nonan");
		Response.Write("|");
		Response.Write(!isFinite(Math.hypot(Infinity, 3)) ? "inf" : "finite");
		Response.Write("|");
		Response.Write(Math.fround(1.337) !== 1.337 ? "diff" : "same");
	`))
	if err != nil {
		t.Fatal(err)
	}
	if out != "0|0|0.549|1.718282|0.693147|3|3|5|-5|31|32|nan|inf|diff" {
		t.Errorf("expected '0|0|0.549|1.718282|0.693147|3|3|5|-5|31|32|nan|inf|diff', got %q", out)
	}
}

func TestJScriptSymbolPrimitiveUniqueness(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		var a = Symbol("x");
		var b = Symbol("x");
		Response.Write((a !== b) ? "yes" : "no");
		Response.Write("|");
		Response.Write(typeof a);
	`))
	if err != nil {
		t.Fatal(err)
	}
	if out != "yes|symbol" {
		t.Errorf("expected 'yes|symbol', got %q", out)
	}
}

func TestJScriptSymbolObjectKeyHiddenFromKeys(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		var s = Symbol("id");
		var o = {};
		o[s] = 42;
		o.visible = 1;
		Response.Write(o[s]);
		Response.Write("|");
		Response.Write(Object.keys(o).join(","));
	`))
	if err != nil {
		t.Fatal(err)
	}
	if out != "42|visible" {
		t.Errorf("expected '42|visible', got %q", out)
	}
}

func TestJScriptArrayFromAndOf(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		var fromArr = Array.from({ length: 3, 0: "a", 1: "b", 2: "c" });
		var ofArr = Array.of(7, 8, 9);
		Response.Write(fromArr.join("-"));
		Response.Write("|");
		Response.Write(ofArr.join("-"));
	`))
	if err != nil {
		t.Fatal(err)
	}
	if out != "a-b-c|7-8-9" {
		t.Errorf("expected 'a-b-c|7-8-9', got %q", out)
	}
}

func TestJScriptRestParameters(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		function pack(first, ...rest) {
			return first + ":" + rest.length + ":" + rest[0] + ":" + rest[1];
		}
		Response.Write(pack("h", 10, 20));
	`))
	if err != nil {
		t.Fatal(err)
	}
	if out != "h:2:10:20" {
		t.Errorf("expected 'h:2:10:20', got %q", out)
	}
}

func TestJScriptSetAndMapBasics(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		var s = new Set();
		s.add("a").add("b");
		var setSize = s.size;
		var before = s.has("a") && s.has("b");
		var deleted = s.delete("a");
		var after = s.has("a");
		s.clear();
		var cleared = s.has("b");

		var m = new Map();
		m.set("k1", 100).set("k2", 200);
		var mBefore = m.has("k1") && m.has("k2");
		var mDeleted = m.delete("k1");
		var mAfter = m.has("k1");
		m.clear();
		var mCleared = m.has("k2");

		Response.Write((setSize === 2 && before && deleted && !after && !cleared) ? "yes" : "no");
		Response.Write("|");
		Response.Write((mBefore && mDeleted && !mAfter && !mCleared) ? "yes" : "no");
	`))
	if err != nil {
		t.Fatal(err)
	}
	if out != "yes|yes" {
		t.Errorf("expected 'yes|yes', got %q", out)
	}
}

func TestJScriptBigIntMixedTypeErrorMessage(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		try {
			Response.Write(10n + 5);
		} catch (e) {
			Response.Write(e.message);
		}
	`))
	if err != nil {
		t.Fatal(err)
	}
	if out != "Cannot mix BigInt and other types, use explicit conversions" {
		t.Errorf("expected BigInt mix error message, got %q", out)
	}
}

func TestJScriptSetForOfIteratesValues(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		var s = new Set();
		s.add("a"); s.add("b"); s.add("c");
		var setResult = "";
		for (const v of s) {
			setResult += v;
		}
		Response.Write(setResult.length);
	`))
	if err != nil {
		t.Fatal(err)
	}
	if out != "3" {
		t.Errorf("expected '3', got %q", out)
	}
}

func TestJScriptSetIterableInitialization(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		var src = {
			[Symbol.iterator]: function() {
				var index = 0;
				return {
					next: function() {
						index = index + 1;
						if (index === 1) {
							return { value: "a", done: false };
						}
						if (index === 2) {
							return { value: "b", done: false };
						}
						return { value: undefined, done: true };
					}
				};
			}
		};
		var s = new Set(src);
		Response.Write(s.has("a") + "|" + s.has("b"));
	`))
	if err != nil {
		t.Fatal(err)
	}
	if out != "True|True" {
		t.Errorf("expected 'True|True', got %q", out)
	}
}

func TestJScriptMapIterableInitialization(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		var src = {
			[Symbol.iterator]: function() {
				var index = 0;
				return {
					next: function() {
						index = index + 1;
						if (index === 1) {
							return { value: ["k1", 10], done: false };
						}
						if (index === 2) {
							return { value: ["k2", 20], done: false };
						}
						return { value: undefined, done: true };
					}
				};
			}
		};
		var m = new Map(src);
		Response.Write(m.get("k1") + "|" + m.get("k2"));
	`))
	if err != nil {
		t.Fatal(err)
	}
	if out != "10|20" {
		t.Errorf("expected '10|20', got %q", out)
	}
}

func TestJScriptSetDetachedPrototypeMethods(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		var s = new Set();
		var add = Set.prototype.add;
		var has = Set.prototype.has;
		add.call(s, "ok");
		Response.Write(has.call(s, "ok"));
	`))
	if err != nil {
		t.Fatal(err)
	}
	if out != "True" {
		t.Errorf("expected 'True', got %q", out)
	}

	_, err = runJScript2(t, jscriptSrc(`
		var has = Set.prototype.has;
		has.call({}, "ok");
	`))
	if err == nil {
		t.Fatal("expected TypeError for Set.prototype.has with incompatible receiver, got nil")
	}
}

func TestJScriptMapDetachedPrototypeMethods(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		var m = new Map();
		var set = Map.prototype.set;
		var get = Map.prototype.get;
		set.call(m, "k", "ok");
		Response.Write(get.call(m, "k"));
	`))
	if err != nil {
		t.Fatal(err)
	}
	if out != "ok" {
		t.Errorf("expected 'ok', got %q", out)
	}

	_, err = runJScript2(t, jscriptSrc(`
		var get = Map.prototype.get;
		get.call({}, "k");
	`))
	if err == nil {
		t.Fatal("expected TypeError for Map.prototype.get with incompatible receiver, got nil")
	}
}

func TestJScriptTDZ(t *testing.T) {
	code := `
		let a = 1;
		{
			let b = a;
			let a = 2; // TDZ error
		}
	`
	_, err := runJScript2(t, jscriptSrc(code))
	if err == nil {
		t.Errorf("Expected ReferenceError for TDZ, got no error")
	} else if !strings.Contains(err.Error(), "Cannot access 'a' before initialization") {
		t.Errorf("Expected ReferenceError for TDZ, got: %v", err)
	}

	code2 := `
		let a = 1;
		{
			const b = a;
			const a = 2; // TDZ error
		}
	`
	_, err = runJScript2(t, jscriptSrc(code2))
	if err == nil {
		t.Errorf("Expected ReferenceError for TDZ const, got no error")
	} else if !strings.Contains(err.Error(), "Cannot access 'a' before initialization") {
		t.Errorf("Expected ReferenceError for TDZ const, got: %v", err)
	}
}

func TestJScriptConstReassignmentTopLevelThrowsTypeError(t *testing.T) {
	_, err := runJScript2(t, jscriptSrc(`
		const PI = 3.14;
		PI = 3.15;
	`))
	if err == nil {
		t.Fatal("expected TypeError for top-level const reassignment, got nil")
	}
	if !strings.Contains(err.Error(), "constant variable") {
		t.Fatalf("expected const reassignment TypeError, got: %v", err)
	}
}

func TestJScriptUnicodeCodePointEscape(t *testing.T) {
	code := `
		var s = "\u{1D306}";
		Response.Write(s.length);
	`
	out, err := runJScript2(t, jscriptSrc(code))
	if err != nil {
		t.Fatal(err)
	}
	if out != "2" { // Surrogate pair
		t.Errorf("expected '2', got %q", out)
	}
}

func TestJScriptRegExpUnicodeFlag(t *testing.T) {
	code := `
		var re = /^\u{1D306}$/u;
		var s = "\u{1D306}";
		Response.Write(re.test(s) ? "pass" : "fail");
	`
	out, err := runJScript2(t, jscriptSrc(code))
	if err != nil {
		t.Fatal(err)
	}
	if out != "pass" {
		t.Errorf("expected 'pass', got %q", out)
	}
}

// ---------------------------------------------------------------------------
// Phase 2: Modern Syntax & Operators
// ---------------------------------------------------------------------------

func TestJScriptPhase2OptionalChaining(t *testing.T) {
	tests := []struct {
		name     string
		source   string
		expected string
	}{
		{"Basic property access", `var a = {b: 1}; Response.Write(a?.b)`, "1"},
		{"Null base", `var a = null; Response.Write(a?.b)`, "undefined"},
		{"Undefined base", `var a; Response.Write(a?.b)`, "undefined"},
		{"Nested property access", `var a = {b: {c: 2}}; Response.Write(a?.b?.c)`, "2"},
		{"Nested null", `var a = {b: null}; Response.Write(a?.b?.c)`, "undefined"},
		{"Call exists", `var a = {b: function() { return 3; }}; Response.Write(a?.b())`, "3"},
		{"Call null base", `var a = null; Response.Write(a?.())`, "undefined"},
		{"Bracket access", `var a = {b: 4}; Response.Write(a?.['b'])`, "4"},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			out, err := runJScript2(t, jscriptSrc(tt.source))
			if err != nil {
				t.Fatalf("failed: %v", err)
			}
			if out != tt.expected {
				t.Errorf("expected %q, got %q", tt.expected, out)
			}
		})
	}
}

func TestJScriptPhase2NullishCoalescing(t *testing.T) {
	tests := []struct {
		name     string
		source   string
		expected string
	}{
		{"Null LHS", `Response.Write(null ?? 1)`, "1"},
		{"Undefined LHS", `Response.Write(undefined ?? 2)`, "2"},
		{"False LHS", `Response.Write(false ?? 3)`, "False"},
		{"Zero LHS", `Response.Write(0 ?? 4)`, "0"},
		{"Empty string LHS", `Response.Write("" ?? 5)`, ""},
		{"Non-nullish LHS", `Response.Write(6 ?? 7)`, "6"},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			out, err := runJScript2(t, jscriptSrc(tt.source))
			if err != nil {
				t.Fatalf("failed: %v", err)
			}
			if out != tt.expected {
				t.Errorf("expected %q, got %q", tt.expected, out)
			}
		})
	}
}

func TestJScriptPhase2LogicalAssignment(t *testing.T) {
	tests := []struct {
		name     string
		source   string
		expected string
	}{
		{"OR assign falsy", `var a = 0; a ||= 1; Response.Write(a)`, "1"},
		{"OR assign truthy", `var a = 2; a ||= 3; Response.Write(a)`, "2"},
		{"AND assign truthy", `var a = 4; a &&= 5; Response.Write(a)`, "5"},
		{"AND assign falsy", `var a = 0; a &&= 6; Response.Write(a)`, "0"},
		{"Coalesce assign nullish", `var a = null; a ??= 7; Response.Write(a)`, "7"},
		{"Coalesce assign non-nullish", `var a = 8; a ??= 9; Response.Write(a)`, "8"},
		{"Assignment result", `var a = 0; Response.Write(a ||= 1)`, "1"},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			out, err := runJScript2(t, jscriptSrc(tt.source))
			if err != nil {
				t.Fatalf("failed: %v", err)
			}
			if out != tt.expected {
				t.Errorf("expected %q, got %q", tt.expected, out)
			}
		})
	}
}

func TestJScriptPhase2Exponentiation(t *testing.T) {
	tests := []struct {
		name     string
		source   string
		expected string
	}{
		{"2 ** 3", `Response.Write(2 ** 3)`, "8"},
		{"3 ** 2", `Response.Write(3 ** 2)`, "9"},
		{"**= operator", `var a = 2; a **= 4; Response.Write(a)`, "16"},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			out, err := runJScript2(t, jscriptSrc(tt.source))
			if err != nil {
				t.Fatalf("failed: %v", err)
			}
			if out != tt.expected {
				t.Errorf("expected %q, got %q", tt.expected, out)
			}
		})
	}
}

func TestJScriptPhase2BigInt(t *testing.T) {
	tests := []struct {
		name     string
		source   string
		expected string
	}{
		{"Literal", `Response.Write(10n)`, "10"},
		{"Addition", `Response.Write(10n + 20n)`, "30"},
		{"Subtraction", `Response.Write(30n - 10n)`, "20"},
		{"Multiplication", `Response.Write(5n * 6n)`, "30"},
		{"Division", `Response.Write(10n / 3n)`, "3"},
		{"Modulo", `Response.Write(10n % 3n)`, "1"},
		{"Exponentiation", `Response.Write(2n ** 10n)`, "1024"},
		{"Negation", `Response.Write(-10n)`, "-10"},
		{"Strict Equality", `Response.Write(10n === 10n)`, "True"},
		{"Strict Inequality", `Response.Write(10n === 10)`, "False"},
		{"Truthy", `if (10n) { Response.Write(1); } else { Response.Write(0); }`, "1"},
		{"Falsy", `if (0n) { Response.Write(1); } else { Response.Write(0); }`, "0"},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			out, err := runJScript2(t, jscriptSrc(tt.source))
			if err != nil {
				t.Fatalf("failed: %v", err)
			}
			if out != tt.expected {
				t.Errorf("expected %q, got %q", tt.expected, out)
			}
		})
	}
}

// ---------------------------------------------------------------------------
// ES6 Binary Data — ArrayBuffer, TypedArrays, DataView
// ---------------------------------------------------------------------------

func TestJScriptArrayBufferCreation(t *testing.T) {
	tests := []struct {
		name     string
		source   string
		expected string
	}{
		{"byteLength", `var b = new ArrayBuffer(8); Response.Write(b.byteLength);`, "8"},
		{"zero length", `var b = new ArrayBuffer(0); Response.Write(b.byteLength);`, "0"},
		{"isView typed array", `
			var b = new ArrayBuffer(4);
			var u = new Uint8Array(b);
			Response.Write(ArrayBuffer.isView(u));
		`, "True"},
		{"isView plain object", `Response.Write(ArrayBuffer.isView({}));`, "False"},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			out, err := runJScript2(t, jscriptSrc(tt.source))
			if err != nil {
				t.Fatalf("failed: %v", err)
			}
			if strings.TrimSpace(out) != tt.expected {
				t.Errorf("expected %q, got %q", tt.expected, strings.TrimSpace(out))
			}
		})
	}
}

func TestJScriptUint8ArrayReadWrite(t *testing.T) {
	tests := []struct {
		name     string
		source   string
		expected string
	}{
		{"length from size", `var a = new Uint8Array(4); Response.Write(a.length);`, "4"},
		{"initial zero", `var a = new Uint8Array(4); Response.Write(a[0]);`, "0"},
		{"write and read", `var a = new Uint8Array(4); a[0] = 42; Response.Write(a[0]);`, "42"},
		{"clamp above 255", `var a = new Uint8ClampedArray(2); a[0] = 300; Response.Write(a[0]);`, "255"},
		{"clamp below 0", `var a = new Uint8ClampedArray(2); a[0] = -5; Response.Write(a[0]);`, "0"},
		{"from array", `var a = new Uint8Array([10,20,30]); Response.Write(a[1]);`, "20"},
		{"byteLength", `var a = new Uint8Array(8); Response.Write(a.byteLength);`, "8"},
		{"byteOffset", `var a = new Uint8Array(8); Response.Write(a.byteOffset);`, "0"},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			out, err := runJScript2(t, jscriptSrc(tt.source))
			if err != nil {
				t.Fatalf("failed: %v", err)
			}
			if strings.TrimSpace(out) != tt.expected {
				t.Errorf("expected %q, got %q", tt.expected, strings.TrimSpace(out))
			}
		})
	}
}

func TestJScriptInt32ArrayReadWrite(t *testing.T) {
	tests := []struct {
		name     string
		source   string
		expected string
	}{
		{"write and read", `var a = new Int32Array(4); a[2] = -1234567; Response.Write(a[2]);`, "-1234567"},
		{"length", `var a = new Int32Array(3); Response.Write(a.length);`, "3"},
		{"byteLength", `var a = new Int32Array(3); Response.Write(a.byteLength);`, "12"},
		{"from array", `var a = new Int32Array([100, 200, 300]); Response.Write(a[2]);`, "300"},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			out, err := runJScript2(t, jscriptSrc(tt.source))
			if err != nil {
				t.Fatalf("failed: %v", err)
			}
			if strings.TrimSpace(out) != tt.expected {
				t.Errorf("expected %q, got %q", tt.expected, strings.TrimSpace(out))
			}
		})
	}
}

func TestJScriptFloat64ArrayReadWrite(t *testing.T) {
	tests := []struct {
		name     string
		source   string
		expected string
	}{
		{"write and read", `var a = new Float64Array(2); a[0] = 3.14; Response.Write(a[0]);`, "3.14"},
		{"from array", `var a = new Float64Array([1.5, 2.5]); Response.Write(a[0] + a[1]);`, "4"},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			out, err := runJScript2(t, jscriptSrc(tt.source))
			if err != nil {
				t.Fatalf("failed: %v", err)
			}
			if strings.TrimSpace(out) != tt.expected {
				t.Errorf("expected %q, got %q", tt.expected, strings.TrimSpace(out))
			}
		})
	}
}

func TestJScriptDataViewGetSet(t *testing.T) {
	tests := []struct {
		name     string
		source   string
		expected string
	}{
		{"getInt8/setInt8", `
			var b = new ArrayBuffer(4);
			var dv = new DataView(b);
			dv.setInt8(0, -42);
			Response.Write(dv.getInt8(0));
		`, "-42"},
		{"getUint8/setUint8", `
			var b = new ArrayBuffer(4);
			var dv = new DataView(b);
			dv.setUint8(1, 200);
			Response.Write(dv.getUint8(1));
		`, "200"},
		{"getInt16 little-endian", `
			var b = new ArrayBuffer(4);
			var dv = new DataView(b);
			dv.setInt16(0, 0x0102, true);
			Response.Write(dv.getInt16(0, true));
		`, "258"},
		{"getInt32 big-endian", `
			var b = new ArrayBuffer(8);
			var dv = new DataView(b);
			dv.setInt32(0, 12345678, false);
			Response.Write(dv.getInt32(0, false));
		`, "12345678"},
		{"getInt32 big-endian sign wrap", `
			var b = new ArrayBuffer(8);
			var dv = new DataView(b);
			dv.setInt32(0, 0xDEADBEEF, false);
			Response.Write(dv.getInt32(0, false));
		`, "-559038737"},
		{"getFloat32", `
			var b = new ArrayBuffer(8);
			var dv = new DataView(b);
			dv.setFloat32(0, 1.5, true);
			var v = dv.getFloat32(0, true);
			Response.Write(v > 1.4 && v < 1.6);
		`, "True"},
		{"getFloat64", `
			var b = new ArrayBuffer(16);
			var dv = new DataView(b);
			dv.setFloat64(0, 2.718281828, true);
			var v = dv.getFloat64(0, true);
			Response.Write(v > 2.71 && v < 2.72);
		`, "True"},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			out, err := runJScript2(t, jscriptSrc(tt.source))
			if err != nil {
				t.Fatalf("failed: %v", err)
			}
			if strings.TrimSpace(out) != tt.expected {
				t.Errorf("expected %q, got %q", tt.expected, strings.TrimSpace(out))
			}
		})
	}
}

func TestJScriptTypedArraySet(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		var a = new Uint8Array(4);
		a.set([10, 20, 30, 40]);
		Response.Write(a[0] + "," + a[1] + "," + a[2] + "," + a[3]);
	`))
	if err != nil {
		t.Fatal(err)
	}
	if strings.TrimSpace(out) != "10,20,30,40" {
		t.Errorf("expected '10,20,30,40', got %q", out)
	}
}

func TestJScriptTypedArraySubarray(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		var a = new Uint8Array([1,2,3,4,5]);
		var b = a.subarray(1, 3);
		Response.Write(b.length + "," + b[0] + "," + b[1]);
	`))
	if err != nil {
		t.Fatal(err)
	}
	if strings.TrimSpace(out) != "2,2,3" {
		t.Errorf("expected '2,2,3', got %q", out)
	}
}

func TestJScriptTypedArrayFill(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		var a = new Uint8Array(4);
		a.fill(7);
		Response.Write(a[0] + "," + a[1] + "," + a[2] + "," + a[3]);
	`))
	if err != nil {
		t.Fatal(err)
	}
	if strings.TrimSpace(out) != "7,7,7,7" {
		t.Errorf("expected '7,7,7,7', got %q", out)
	}
}

func TestJScriptArrayBufferSlice(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		var b = new ArrayBuffer(8);
		var dv = new DataView(b);
		dv.setUint8(0, 1);
		dv.setUint8(1, 2);
		dv.setUint8(2, 3);
		var b2 = b.slice(1, 3);
		var dv2 = new DataView(b2);
		Response.Write(b2.byteLength + "," + dv2.getUint8(0) + "," + dv2.getUint8(1));
	`))
	if err != nil {
		t.Fatal(err)
	}
	if strings.TrimSpace(out) != "2,2,3" {
		t.Errorf("expected '2,2,3', got %q", out)
	}
}

func TestJScriptTypedArrayForOf(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		var a = new Uint8Array([10, 20, 30]);
		var sum = 0;
		for (var v of a) { sum += v; }
		Response.Write(sum);
	`))
	if err != nil {
		t.Fatal(err)
	}
	if strings.TrimSpace(out) != "60" {
		t.Errorf("expected '60', got %q", out)
	}
}

// ---------------------------------------------------------------------------
// ES6 Well-Known Symbols
// ---------------------------------------------------------------------------

func TestJScriptWellKnownSymbolIterator(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		Response.Write(typeof Symbol.iterator === "symbol");
	`))
	if err != nil {
		t.Fatal(err)
	}
	if strings.TrimSpace(out) != "True" {
		t.Errorf("expected 'True', got %q", out)
	}
}

func TestJScriptWellKnownSymbolToStringTag(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		Response.Write(typeof Symbol.toStringTag === "symbol");
	`))
	if err != nil {
		t.Fatal(err)
	}
	if strings.TrimSpace(out) != "True" {
		t.Errorf("expected 'True', got %q", out)
	}
}

func TestJScriptWellKnownSymbolSpecies(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		Response.Write(typeof Symbol.species === "symbol");
	`))
	if err != nil {
		t.Fatal(err)
	}
	if strings.TrimSpace(out) != "True" {
		t.Errorf("expected 'True', got %q", out)
	}
}

func TestJScriptWellKnownSymbolHasInstance(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		Response.Write(typeof Symbol.hasInstance === "symbol");
	`))
	if err != nil {
		t.Fatal(err)
	}
	if strings.TrimSpace(out) != "True" {
		t.Errorf("expected 'True', got %q", out)
	}
}

func TestJScriptWellKnownSymbolToPrimitive(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		Response.Write(typeof Symbol.toPrimitive === "symbol");
	`))
	if err != nil {
		t.Fatal(err)
	}
	if strings.TrimSpace(out) != "True" {
		t.Errorf("expected 'True', got %q", out)
	}
}

func TestJScriptSymbolFor(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		var s1 = Symbol.for("myKey");
		var s2 = Symbol.for("myKey");
		Response.Write(s1 === s2);
	`))
	if err != nil {
		t.Fatal(err)
	}
	if strings.TrimSpace(out) != "True" {
		t.Errorf("expected 'True', got %q", out)
	}
}

func TestJScriptSymbolKeyFor(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		var s = Symbol.for("testKey");
		Response.Write(Symbol.keyFor(s));
	`))
	if err != nil {
		t.Fatal(err)
	}
	if strings.TrimSpace(out) != "testKey" {
		t.Errorf("expected 'testKey', got %q", out)
	}
}

func TestJScriptSymbolKeyForUnregistered(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		var s = Symbol("notRegistered");
		var k = Symbol.keyFor(s);
		Response.Write(k === undefined);
	`))
	if err != nil {
		t.Fatal(err)
	}
	if strings.TrimSpace(out) != "True" {
		t.Errorf("expected 'True', got %q", out)
	}
}

func TestJScriptTailCallDeepRecursion(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		function sum(n, acc) {
			if (n === 0) {
				return acc;
			}
			return sum(n - 1, acc + 1);
		}
		Response.Write(sum(100000, 0));
	`))
	if err != nil {
		t.Fatal(err)
	}
	if out != "100000" {
		t.Errorf("expected '100000', got %q", out)
	}
}

func TestJScriptTailCallInsideTryCatchBypassesTCO(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		function sumInTry(n, acc) {
			try {
				if (n === 0) {
					return acc;
				}
				return sumInTry(n - 1, acc + 1);
			} catch (e) {
				return -1;
			}
		}
		Response.Write(sumInTry(128, 0));
	`))
	if err != nil {
		t.Fatal(err)
	}
	if out != "128" {
		t.Errorf("expected '128', got %q", out)
	}
}

// ---------------------------------------------------------------------------
// ES6 Iteration Protocol (Sub-Phase 5.1)
// ---------------------------------------------------------------------------

func TestJScriptIterationProtocol(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		Response.Write(typeof Array.prototype + "|");
		Response.Write(typeof Array.prototype.values + "|");
		Response.Write(typeof Array.prototype["__js_sym__-1"] + "|");
		var arr = [1, 2, 3];
		var s = Symbol.iterator;
		Response.Write(typeof arr[s]);
	`))
	if err != nil {
		t.Fatal(err)
	}
	if out != "object|function|function|function" {
		t.Errorf("expected 'object|function|function|function', got %q", out)
	}
}

func TestJScriptIterationProtocolFull(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		var arr = [1, 2, 3];
		var it = arr[Symbol.iterator]();
		Response.Write(typeof it + "|");
		var res = it.next();
		Response.Write(typeof res + "|");
		Response.Write(res.value + "|");
		Response.Write(res.done);
	`))
	if err != nil {
		t.Fatal(err)
	}
	if out != "object|object|1|False" {
		t.Errorf("expected 'object|object|1|False', got %q", out)
	}
}

func TestJScriptStringIterationProtocol(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		var s = "ABC";
		var it = s[Symbol.iterator]();
		Response.Write(it.next().value + "|");
		Response.Write(it.next().value + "|");
		Response.Write(it.next().value + "|");
		Response.Write(it.next().done);
	`))
	if err != nil {
		t.Fatal(err)
	}
	if out != "A|B|C|True" {
		t.Errorf("expected 'A|B|C|True', got %q", out)
	}
}

func TestJScriptForOfCustomIterable(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		var myIterable = {
			[Symbol.iterator]: function() {
				var count = 0;
				return {
					next: function() {
						if (count < 3) {
							return { value: ++count, done: false };
						}
						return { value: undefined, done: true };
					}
				};
			}
		};
		var result = "";
		for (var x of myIterable) {
			result += x;
		}
		Response.Write(result);
	`))
	if err != nil {
		t.Fatal(err)
	}
	if out != "123" {
		t.Errorf("expected '123', got %q", out)
	}
}

func TestJScriptForOfArray(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		var arr = [10, 20, 30];
		var result = "";
		for (var x of arr) {
			result += x + "|";
		}
		Response.Write(result);
	`))
	if err != nil {
		t.Fatal(err)
	}
	if out != "10|20|30|" {
		t.Errorf("expected '10|20|30|', got %q", out)
	}
}

// ---------------------------------------------------------------------------
// ES6 Classes (Phase 6)
// ---------------------------------------------------------------------------

func TestJScriptClassInstanceMethods(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		class Calculator {
			constructor(x) {
				this.x = x;
			}
			add(y) {
				return this.x + y;
			}
			multiply(y) {
				return this.x * y;
			}
		}
		var calc = new Calculator(10);
		Response.Write(calc.add(5) + "|" + calc.multiply(3));
	`))
	if err != nil {
		t.Fatal(err)
	}
	if out != "15|30" {
		t.Errorf("expected '15|30', got %q", out)
	}
}

func TestJScriptClassStaticMethods(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		class Utils {
			static square(x) {
				return x * x;
			}
			static cube(x) {
				return x * x * x;
			}
		}
		Response.Write(Utils.square(4) + "|" + Utils.cube(3));
	`))
	if err != nil {
		t.Fatal(err)
	}
	if out != "16|27" {
		t.Errorf("expected '16|27', got %q", out)
	}
}

func TestJScriptClassAccessors(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		class Person {
			constructor(name) {
				this._name = name;
			}
			get name() {
				return this._name.toUpperCase();
			}
			set name(value) {
				this._name = value;
			}
		}
		var p = new Person("alice");
		Response.Write(p.name + "|");
		p.name = "bob";
		Response.Write(p.name);
	`))
	if err != nil {
		t.Fatal(err)
	}
	if out != "ALICE|BOB" {
		t.Errorf("expected 'ALICE|BOB', got %q", out)
	}
}

func TestJScriptClassStrictModeEnforcement(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		class StrictTest {
			constructor() {
				// In strict mode, assigning to undeclared variable throws ReferenceError
				try {
					undeclared = 1;
				} catch(e) {
					Response.Write("catch");
				}
			}
			method() {
				try {
					undeclared = 2;
				} catch(e) {
					Response.Write("|catch");
				}
			}
		}
		var s = new StrictTest();
		s.method();
	`))
	if err != nil {
		t.Fatal(err)
	}
	if out != "catch|catch" {
		t.Errorf("expected 'catch|catch', got %q", out)
	}
}

func TestJScriptClassInheritance(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		class Animal {}
		class Dog extends Animal {}
		Response.Write("ok");
	`))
	if err != nil {
		t.Fatal(err)
	}
	if out != "ok" {
		t.Errorf("expected 'ok', got %q", out)
	}
}

func TestJScriptClassExtendsNull(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		class NullBase extends null {}
		var instance = new NullBase();
		Response.Write(instance instanceof NullBase);
	`))
	if err != nil {
		t.Fatal(err)
	}
	if out != "True" {
		t.Errorf("expected 'True', got %q", out)
	}
}

func TestJScriptClassExtendsRejectsInvalidSuperclass(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		try {
			class Broken extends 123 {}
			Response.Write("FAIL");
		} catch (e) {
			Response.Write(String(e).indexOf("TypeError") !== -1);
		}
	`))
	if err != nil {
		t.Fatal(err)
	}
	if out != "True" {
		t.Errorf("expected invalid superclass to raise TypeError, got %q", out)
	}
}

func TestJScriptClassStaticInheritance(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		class Base {
			static staticMethod() {
				return "static";
			}
		}
		class Derived extends Base {}
		Response.Write(Derived.staticMethod());
	`))
	if err != nil {
		t.Fatal(err)
	}
	if out != "static" {
		t.Errorf("expected 'static', got %q", out)
	}
}

func TestJScriptClassSuperConstructorTDZ(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		class Base {
			constructor() { this.x = 1; }
		}
		class Derived extends Base {
			constructor() {
				super();
			}
		}
		var d = new Derived();
		Response.Write(d.x);
	`))
	if err != nil {
		t.Fatal(err)
	}
	if out != "1" {
		t.Errorf("expected '1', got %q", out)
	}
}

func TestJScriptClassSuperDelegation(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		class Base {
			foo() { return "base"; }
			static staticFoo() { return "static base"; }
		}
		class Derived extends Base {
			foo() {
				return super.foo() + " derived";
			}
			static staticFoo() {
				return super.staticFoo() + " static derived";
			}
			setX(val) {
				super.x = val;
			}
		}
		var d = new Derived();
		d.setX(42);
		Response.Write(d.foo() + " | " + Derived.staticFoo() + " | x=" + d.x);
	`))
	if err != nil {
		t.Fatal(err)
	}
	if out != "base derived | static base static derived | x=42" {
		t.Errorf("expected 'base derived | static base static derived | x=42', got %q", out)
	}
}

func TestJScriptClassFields(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		class Base {
			x = 1;
			static s = "static";
		}
		class Derived extends Base {
			y = 2;
			constructor() {
				super();
				this.z = 3;
			}
		}
		var d = new Derived();
		Response.Write("x=" + d.x + ", y=" + d.y + ", z=" + d.z + ", s=" + Base.s);
	`))
	if err != nil {
		t.Fatal(err)
	}
	if out != "x=1, y=2, z=3, s=static" {
		t.Errorf("expected 'x=1, y=2, z=3, s=static', got %q", out)
	}
}

// ---------------------------------------------------------------------------
// Symbol-Based Keys for Maps/WeakMaps (Phase 3.1)
// ---------------------------------------------------------------------------

func TestJScriptWeakMapSymbolKey(t *testing.T) {
	out := runASPSourceForTest(t, `<script runat="server" language="JScript">
		var wm = new WeakMap();
		var key = Symbol("foo");
		wm.set(key, "bar");
		Response.Write(wm.get(key));
	</script>`)
	if out != "bar" {
		t.Errorf("expected 'bar', got %q", out)
	}
}

func TestJScriptMapSymbolKey(t *testing.T) {
	out := runASPSourceForTest(t, `<script runat="server" language="JScript">
		var m = new Map();
		var key = Symbol("foo");
		m.set(key, "bar");
		Response.Write(m.get(key));
	</script>`)
	if out != "bar" {
		t.Errorf("expected 'bar', got %q", out)
	}
}

func TestJScriptWeakSetSymbolKey(t *testing.T) {
	out := runASPSourceForTest(t, `<script runat="server" language="JScript">
		var ws = new WeakSet();
		var key = Symbol("foo");
		ws.add(key);
		Response.Write(ws.has(key));
	</script>`)
	if out != "True" {
		t.Errorf("expected 'True', got %q", out)
	}
}

func TestJScriptWeakMapInvalidSymbolKeys(t *testing.T) {
	// Registered symbol
	_, err := runASPSourceForTestWithErr(t, `<script runat="server" language="JScript">
		var wm = new WeakMap();
		var key = Symbol.for("foo");
		wm.set(key, "bar");
	</script>`)
	if err == nil {
		t.Error("expected TypeError for registered symbol in WeakMap.set, got nil")
	}

	// Well-known symbol
	_, err = runASPSourceForTestWithErr(t, `<script runat="server" language="JScript">
		var wm = new WeakMap();
		var key = Symbol.iterator;
		wm.set(key, "bar");
	</script>`)
	if err == nil {
		t.Error("expected TypeError for well-known symbol in WeakMap.set, got nil")
	}
}

func TestJScriptWeakRef(t *testing.T) {
	out := runASPSourceForTest(t, `<script runat="server" language="JScript">
		var obj = { x: 1 };
		var wr = new WeakRef(obj);
		var target = wr.deref();
		Response.Write("wr.deref() is obj: " + (target === obj) + " | ");
		if (target) {
			Response.Write("target.x: " + target.x);
		} else {
			Response.Write("null/undefined");
		}
	</script>`)
	if out != "wr.deref() is obj: True | target.x: 1" {
		t.Errorf("expected 'wr.deref() is obj: True | target.x: 1', got %q", out)
	}
}

func TestJScriptFinalizationRegistry(t *testing.T) {
	out := runASPSourceForTest(t, `<script runat="server" language="JScript">
		var registry = new FinalizationRegistry(function(held) {
			// Callback won't be called in our current VM unless we implement a sweep.
		});
		var obj = {};
		registry.register(obj, "some data", obj);
		registry.unregister(obj);
		Response.Write("ok");
	</script>`)
	if out != "ok" {
		t.Errorf("expected 'ok', got %q", out)
	}
}

func TestJScriptPrivateClassFields(t *testing.T) {
	out := runASPSourceForTest(t, `<script runat="server" language="JScript">
		class Test {
			#x = 10;
			static #y = 20;

			get() {
				return this.#x;
			}
			
			set(val) {
				this.#x = val;
			}

			inc() {
				this.#x++;
				return this.#x;
			}
			
			static getStatic() {
				return Test.#y;
			}

			static setStatic(val) {
				Test.#y = val;
			}
		}

		var t = new Test();
		Response.Write("x: " + t.get() + " | ");
		t.set(15);
		Response.Write("set x: " + t.get() + " | ");
		Response.Write("inc x: " + t.inc() + " | ");
		Response.Write("static y: " + Test.getStatic() + " | ");
		Test.setStatic(25);
		Response.Write("set static y: " + Test.getStatic());
	</script>`)

	expected := "x: 10 | set x: 15 | inc x: 16 | static y: 20 | set static y: 25"
	if out != expected {
		t.Errorf("expected %q, got %q", expected, out)
	}
}

func TestJScriptUsingDisposesOnScopeExit(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		var events = [];
		var resource = {
			[Symbol.dispose]: function() {
				events.push("dispose");
			}
		};
		{
			using r = resource;
			events.push("body");
		}
		Response.Write(events.join(","));
	`))
	if err != nil {
		t.Fatal(err)
	}
	if out != "body,dispose" {
		t.Errorf("expected 'body,dispose', got %q", out)
	}
}

func TestJScriptUsingDisposesOnThrow(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		var resource = {
			[Symbol.dispose]: function() {
				Response.Write("D");
			}
		};
		try {
			{
				using r = resource;
				throw "boom";
			}
		} catch (e) {
			Response.Write("C");
		}
	`))
	if err != nil {
		t.Fatal(err)
	}
	if out != "DC" {
		t.Errorf("expected 'DC', got %q", out)
	}
}

func TestJScriptUsingDisposesInReverseOrder(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		var aRes = { [Symbol.dispose]: function() { Response.Write("A"); } };
		var bRes = { [Symbol.dispose]: function() { Response.Write("B"); } };
		{
			using a = aRes;
			using b = bRes;
		}
	`))
	if err != nil {
		t.Fatal(err)
	}
	if out != "BA" {
		t.Errorf("expected 'BA', got %q", out)
	}
}

func TestJScriptAsyncUsingCallsSymbolAsyncDispose(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		var resource = {
			[Symbol.asyncDispose]: function() {
				Response.Write("A");
			}
		};
		{
			async using r = resource;
			Response.Write("I");
		}
	`))
	if err != nil {
		t.Fatal(err)
	}
	if out != "IA" {
		t.Errorf("expected 'IA', got %q", out)
	}
}

func TestJScriptUsingRequiresInitializer(t *testing.T) {
	_, err := runJScript2(t, jscriptSrc(`
		{
			using r;
		}
	`))
	if err == nil {
		t.Fatal("expected compile error for using declaration without initializer")
	}
}
