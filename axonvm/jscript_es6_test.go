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

func TestJScriptObjectAssignKeysValuesEntries(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		var target = { a: 1 };
		Object.assign(target, { b: 2 }, { c: 3 });
		Response.Write(target.a + "," + target.b + "," + target.c);
		Response.Write("|");
		Response.Write(Object.keys(target).join(","));
		Response.Write("|");
		Response.Write(Object.values(target).join(","));
		Response.Write("|");
		var e = Object.entries(target);
		Response.Write(e[0][0] + ":" + e[0][1] + ";" + e[1][0] + ":" + e[1][1] + ";" + e[2][0] + ":" + e[2][1]);
	`))
	if err != nil {
		t.Fatal(err)
	}
	if out != "1,2,3|a,b,c|1,2,3|a:1;b:2;c:3" {
		t.Errorf("expected '1,2,3|a,b,c|1,2,3|a:1;b:2;c:3', got %q", out)
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

func TestJScriptObjectKeysNullThrowsTypeError(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		var ok = false;
		try {
			Object.keys(null);
		} catch (e) {
			ok = (String(e).indexOf("TypeError") !== -1);
		}
		Response.Write(ok ? "yes" : "no");
	`))
	if err != nil {
		t.Fatal(err)
	}
	if out != "yes" {
		t.Errorf("expected 'yes', got %q", out)
	}
}

func TestJScriptArraySpreadNullThrowsTypeError(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		var ok = false;
		try {
			var arr = [...null];
			Response.Write(arr.length);
		} catch (e) {
			ok = (String(e).indexOf("TypeError") !== -1);
		}
		Response.Write(ok ? "yes" : "no");
	`))
	if err != nil {
		t.Fatal(err)
	}
	if out != "yes" {
		t.Errorf("expected 'yes', got %q", out)
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

func TestJScriptSymbolConstructorThrows(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		var ok = false;
		try {
			var x = new Symbol("x");
		} catch (e) {
			ok = (String(e).indexOf("TypeError") !== -1);
		}
		Response.Write(ok ? "yes" : "no");
	`))
	if err != nil {
		t.Fatal(err)
	}
	if out != "yes" {
		t.Errorf("expected 'yes', got %q", out)
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

// ---------------------------------------------------------------------------
// Phase 1: Computed Property Names
// ---------------------------------------------------------------------------

func TestJScriptComputedPropertyNames(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		var key = "name";
		var o = { [key]: "Alice" };
		Response.Write(o.name);
	`))
	if err != nil {
		t.Fatal(err)
	}
	if out != "Alice" {
		t.Errorf("expected 'Alice', got %q", out)
	}
}

func TestJScriptComputedPropertyNamesExpression(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		var prefix = "greet";
		var o = { [prefix + "_en"]: "Hello", [prefix + "_fr"]: "Bonjour" };
		Response.Write(o.greet_en);
		Response.Write("|");
		Response.Write(o.greet_fr);
	`))
	if err != nil {
		t.Fatal(err)
	}
	if out != "Hello|Bonjour" {
		t.Errorf("expected 'Hello|Bonjour', got %q", out)
	}
}

func TestJScriptComputedPropertyNamesMixed(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		var k = "dynamic";
		var o = { static: 1, [k]: 2 };
		Response.Write(o.static);
		Response.Write("|");
		Response.Write(o.dynamic);
	`))
	if err != nil {
		t.Fatal(err)
	}
	if out != "1|2" {
		t.Errorf("expected '1|2', got %q", out)
	}
}

// ---------------------------------------------------------------------------
// Phase 2: For...Of Loops
// ---------------------------------------------------------------------------

func TestJScriptForOfArray(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		var arr = [10, 20, 30];
		var result = "";
		for (var x of arr) {
			result += x + ",";
		}
		Response.Write(result);
	`))
	if err != nil {
		t.Fatal(err)
	}
	if out != "10,20,30," {
		t.Errorf("expected '10,20,30,', got %q", out)
	}
}

func TestJScriptForOfString(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		var result = "";
		for (var ch of "abc") {
			result += ch;
		}
		Response.Write(result);
	`))
	if err != nil {
		t.Fatal(err)
	}
	if out != "abc" {
		t.Errorf("expected 'abc', got %q", out)
	}
}

func TestJScriptForOfLet(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		var sum = 0;
		for (let n of [1, 2, 3, 4]) {
			sum += n;
		}
		Response.Write(sum);
	`))
	if err != nil {
		t.Fatal(err)
	}
	if out != "10" {
		t.Errorf("expected '10', got %q", out)
	}
}

func TestJScriptForOfBreak(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		var result = "";
		for (var x of [1, 2, 3, 4, 5]) {
			if (x === 3) break;
			result += x + ",";
		}
		Response.Write(result);
	`))
	if err != nil {
		t.Fatal(err)
	}
	if out != "1,2," {
		t.Errorf("expected '1,2,', got %q", out)
	}
}

func TestJScriptForOfContinue(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		var result = "";
		for (var x of [1, 2, 3, 4]) {
			if (x === 2) continue;
			result += x + ",";
		}
		Response.Write(result);
	`))
	if err != nil {
		t.Fatal(err)
	}
	if out != "1,3,4," {
		t.Errorf("expected '1,3,4,', got %q", out)
	}
}

func TestJScriptForOfEmpty(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		var result = "ok";
		for (var x of []) {
			result = "fail";
		}
		Response.Write(result);
	`))
	if err != nil {
		t.Fatal(err)
	}
	if out != "ok" {
		t.Errorf("expected 'ok', got %q", out)
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
