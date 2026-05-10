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
	"strings"
	"testing"
)

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

func TestJScriptStringEscapeSequences(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		var okNewline = "A\nB".length === 3 && "A\nB".charCodeAt(1) === 10;
		var okSingleQuote = 'It\'s' === "It's";
		var okDoubleQuote = "\"ok\"" === '"ok"';
		var okBackslash = "\\".length === 1 && "\\".charCodeAt(0) === 92;
		var okBackspace = "\b".charCodeAt(0) === 8;
		var okFormFeed = "\f".charCodeAt(0) === 12;
		var okCarriageReturn = "\r".charCodeAt(0) === 13;
		var okTab = "\t".charCodeAt(0) === 9;
		var okVerticalTab = "\v".charCodeAt(0) === 11;
		Response.Write(okNewline + "|" + okSingleQuote + "|" + okDoubleQuote + "|" + okBackslash + "|" + okBackspace + "|" + okFormFeed + "|" + okCarriageReturn + "|" + okTab + "|" + okVerticalTab);
	`))
	if err != nil {
		t.Fatal(err)
	}
	if out != "True|True|True|True|True|True|True|True|True" {
		t.Errorf("unexpected escape sequence behavior: %q", out)
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
		function multiply(a, b) {
			if (b === undefined) b = 2;
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

		Response.Write((before && deleted && !after && !cleared) ? "yes" : "no");
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
