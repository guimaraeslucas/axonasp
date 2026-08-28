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
	"math"
	"testing"

	"g3pix.com.br/axonasp/v2/axonvm/asp"
)

func TestJScriptArrayAt(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		var arr = [1, 2, 3];
		Response.Write(arr.at(0) + ",");
		Response.Write(arr.at(1) + ",");
		Response.Write(arr.at(-1) + ",");
		Response.Write(arr.at(5) + ",");
		Response.Write("hello".at(0) + ",");
		Response.Write("hello".at(-1));
	`))
	if err != nil {
		t.Fatal(err)
	}
	expected := "1,2,3,undefined,h,o"
	if out != expected {
		t.Errorf("expected %q, got %q", expected, out)
	}
}

func TestJScriptArrayFlat(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		var arr = [1, [2, [3, 4]]];
		Response.Write(JSON.stringify(arr.flat()) + "|");
		Response.Write(JSON.stringify(arr.flat(2)));
	`))
	if err != nil {
		t.Fatal(err)
	}
	expected := "[1,2,[3,4]]|[1,2,3,4]"
	if out != expected {
		t.Errorf("expected %q, got %q", expected, out)
	}
}

func TestJScriptArrayFlatMap(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		var arr = [1, 2, 3];
		var res = arr.flatMap(x => [x, x * 2]);
		Response.Write(JSON.stringify(res));
	`))
	if err != nil {
		t.Fatal(err)
	}
	expected := "[1,2,2,4,3,6]"
	if out != expected {
		t.Errorf("expected %q, got %q", expected, out)
	}
}

func TestJScriptArrayImmutable(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		var arr = [3, 1, 2];
		var sorted = arr.toSorted();
		var reversed = arr.toReversed();
		var spliced = arr.toSpliced(1, 1, 4, 5);
		Response.Write(JSON.stringify(arr) + "|");
		Response.Write(JSON.stringify(sorted) + "|");
		Response.Write(JSON.stringify(reversed) + "|");
		Response.Write(JSON.stringify(spliced));
	`))
	if err != nil {
		t.Fatal(err)
	}
	expected := "[3,1,2]|[1,2,3]|[2,1,3]|[3,4,5,2]"
	if out != expected {
		t.Errorf("expected %q, got %q", expected, out)
	}
}

func TestJScriptObjectFromEntries(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		var entries = [["a", 1], ["b", 2]];
		var obj = Object.fromEntries(entries);
		Response.Write(obj.a + "," + obj.b);
	`))
	if err != nil {
		t.Fatal(err)
	}
	expected := "1,2"
	if out != expected {
		t.Errorf("expected %q, got %q", expected, out)
	}
}

// Phase 1: Type & Prototype Fixes

func TestJScriptTypeOfDateCall(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		Response.Write(typeof Date());
	`))
	if err != nil {
		t.Fatal(err)
	}
	expected := "string"
	if out != expected {
		t.Errorf("expected %q, got %q", expected, out)
	}
}

func TestJScriptTypeOfNewDate(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		Response.Write(typeof new Date());
	`))
	if err != nil {
		t.Fatal(err)
	}
	expected := "object"
	if out != expected {
		t.Errorf("expected %q, got %q", expected, out)
	}
}

func TestJScriptTypeOfMath(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		Response.Write(typeof Math);
	`))
	if err != nil {
		t.Fatal(err)
	}
	expected := "object"
	if out != expected {
		t.Errorf("expected %q, got %q", expected, out)
	}
}

func TestJScriptTypeOfParseInt(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		Response.Write(typeof parseInt);
	`))
	if err != nil {
		t.Fatal(err)
	}
	expected := "function"
	if out != expected {
		t.Errorf("expected %q, got %q", expected, out)
	}
}

func TestJScriptInstanceOfObjectLiteral(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		Response.Write(({} instanceof Object) ? "true" : "false");
	`))
	if err != nil {
		t.Fatal(err)
	}
	expected := "true"
	if out != expected {
		t.Errorf("expected %q, got %q", expected, out)
	}
}

func TestJScriptInstanceOfNewDateObject(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		Response.Write((new Date() instanceof Object) ? "true" : "false");
	`))
	if err != nil {
		t.Fatal(err)
	}
	expected := "true"
	if out != expected {
		t.Errorf("expected %q, got %q", expected, out)
	}
}

func TestJScriptInstanceOfRegExpLiteral(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		Response.Write((/abc/ instanceof RegExp) ? "true" : "false");
	`))
	if err != nil {
		t.Fatal(err)
	}
	expected := "true"
	if out != expected {
		t.Errorf("expected %q, got %q", expected, out)
	}
}

func TestJScriptInstanceOfDateConstructor(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		Response.Write((Date instanceof Object) ? "true" : "false");
	`))
	if err != nil {
		t.Fatal(err)
	}
	expected := "true"
	if out != expected {
		t.Errorf("expected %q, got %q", expected, out)
	}
}

func TestJScriptTypeOfMathMax(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		Response.Write(typeof Math.max);
	`))
	if err != nil {
		t.Fatal(err)
	}
	expected := "function"
	if out != expected {
		t.Errorf("expected %q, got %q", expected, out)
	}
}

func TestJScriptTypeOfEval(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		Response.Write(typeof eval);
	`))
	if err != nil {
		t.Fatal(err)
	}
	expected := "function"
	if out != expected {
		t.Errorf("expected %q, got %q", expected, out)
	}
}

func TestJScriptInstanceOfDate(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		Response.Write((new Date() instanceof Date) ? "true" : "false");
	`))
	if err != nil {
		t.Fatal(err)
	}
	expected := "true"
	if out != expected {
		t.Errorf("expected %q, got %q", expected, out)
	}
}

// Phase 2: Unary Operators & Type Coercion

func TestJScriptUnaryPlusEmptyString(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		Response.Write(+"" === 0 ? "true" : "false");
	`))
	if err != nil {
		t.Fatal(err)
	}
	expected := "true"
	if out != expected {
		t.Errorf("expected %q, got %q", expected, out)
	}
}

func TestJScriptUnaryPlusBool(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		Response.Write((+true === 1 && +false === 0) ? "true" : "false");
	`))
	if err != nil {
		t.Fatal(err)
	}
	expected := "true"
	if out != expected {
		t.Errorf("expected %q, got %q", expected, out)
	}
}

func TestJScriptUnaryPlusArray(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		Response.Write((+[1] === 1 && +[] === 0) ? "true" : "false");
	`))
	if err != nil {
		t.Fatal(err)
	}
	expected := "true"
	if out != expected {
		t.Errorf("expected %q, got %q", expected, out)
	}
}

func TestJScriptArrayJoinNullUndefined(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		Response.Write([1, null, 3].join("-") + "|" + [1, void 0, 3].join("-"));
	`))
	if err != nil {
		t.Fatal(err)
	}
	expected := "1--3|1--3"
	if out != expected {
		t.Errorf("expected %q, got %q", expected, out)
	}
}

func TestJScriptUnaryPlusNumber(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		Response.Write((+"123" === 123 && +"-5" === -5 && +" " === 0) ? "true" : "false");
	`))
	if err != nil {
		t.Fatal(err)
	}
	expected := "true"
	if out != expected {
		t.Errorf("expected %q, got %q", expected, out)
	}
}

// Phase 3: Object.prototype.toString & Math Quirks

func TestJScriptObjectToStringStringWrapper(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		Response.Write(Object.prototype.toString.call(new String("x")));
	`))
	if err != nil {
		t.Fatal(err)
	}
	expected := "[object String]"
	if out != expected {
		t.Errorf("expected %q, got %q", expected, out)
	}
}

func TestJScriptObjectToStringNumberWrapper(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		Response.Write(Object.prototype.toString.call(new Number(1)));
	`))
	if err != nil {
		t.Fatal(err)
	}
	expected := "[object Number]"
	if out != expected {
		t.Errorf("expected %q, got %q", expected, out)
	}
}

func TestJScriptObjectToStringBooleanWrapper(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		Response.Write(Object.prototype.toString.call(new Boolean(false)));
	`))
	if err != nil {
		t.Fatal(err)
	}
	expected := "[object Boolean]"
	if out != expected {
		t.Errorf("expected %q, got %q", expected, out)
	}
}

// TestJScriptStringTypeConversion verifies that the JScript String() function
// performs proper type conversion (ES5 §15.5.1) rather than the VBScript
// String(n, char) builtin which repeats the first character n times.
// Regression test: the JScript root environment must be initialized before
// the first global name lookup so that builtins like String, Number, Boolean,
// Date, Array resolve to their JScript intrinsic objects instead of VTBuiltin.
func TestJScriptStringTypeConversion(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		Response.Write(String(42) + "|");
		Response.Write(String(true) + "|");
		Response.Write(String(null) + "|");
		Response.Write(String(undefined) + "|");
		Response.Write(String(3.14) + "|");
		Response.Write(String("hello") + "|");
		Response.Write(String(0) + "|");
		Response.Write(String(false) + "|");
		Response.Write(String(NaN));
	`))
	if err != nil {
		t.Fatal(err)
	}
	expected := "42|true|null|undefined|3.14|hello|0|false|NaN"
	if out != expected {
		t.Errorf("expected %q, got %q", expected, out)
	}
}

// TestJScriptStringTypeConversionNoArgs verifies String() with zero arguments
// returns empty string (ES5 behavior), not VBScript String() which would
// also return empty but for different reasons.
func TestJScriptStringTypeConversionNoArgs(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		Response.Write("[" + String() + "]");
	`))
	if err != nil {
		t.Fatal(err)
	}
	expected := "[]"
	if out != expected {
		t.Errorf("expected %q, got %q", expected, out)
	}
}

// TestJScriptNumberTypeConversion verifies Number() works in JScript mode
// and does not resolve to any VBScript builtin.
func TestJScriptNumberTypeConversion(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		Response.Write(Number("42") + "|");
		Response.Write(Number("3.14") + "|");
		Response.Write(Number("") + "|");
		Response.Write(Number(true) + "|");
		Response.Write(Number(false));
	`))
	if err != nil {
		t.Fatal(err)
	}
	expected := "42|3.14|0|1|0"
	if out != expected {
		t.Errorf("expected %q, got %q", expected, out)
	}
}

// TestJScriptBooleanTypeConversion verifies Boolean() works in JScript mode.
func TestJScriptBooleanTypeConversion(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		Response.Write((Boolean(1) ? "true" : "false") + "|");
		Response.Write((Boolean(0) ? "true" : "false") + "|");
		Response.Write((Boolean("") ? "true" : "false") + "|");
		Response.Write((Boolean("hello") ? "true" : "false") + "|");
		Response.Write((Boolean(null) ? "true" : "false") + "|");
		Response.Write((Boolean(undefined) ? "true" : "false"));
	`))
	if err != nil {
		t.Fatal(err)
	}
	expected := "true|false|false|true|false|false"
	if out != expected {
		t.Errorf("expected %q, got %q", expected, out)
	}
}

// TestJScriptStringBuiltinNotVBScript verifies that the VBScript String(n, char)
// builtin is NOT called when using String() in JScript mode. In VBScript,
// String(3, "x") returns "xxx". In JScript, String(3, "x") should still call
// String() with 2 args — which per ES5 spec ignores extra args and returns
// String(3) = "3".
func TestJScriptStringBuiltinNotVBScript(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		Response.Write(String(3, "x"));
	`))
	if err != nil {
		t.Fatal(err)
	}
	// JScript String() ignores extra args → returns String(3) = "3"
	// VBScript String(3, "x") would return "xxx"
	expected := "3"
	if out != expected {
		t.Errorf("expected %q, got %q (if got 'xxx', VBScript builtin is being called)", expected, out)
	}
}

func TestJScriptMathRoundNegative(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		var r1 = Math.round(-1.5);
		var r2 = Math.round(-1.51);
		var r3 = Math.round(-0.5);
		var r4 = Math.round(-0.1);
		Response.Write(r1 + "," + r2 + "," + (1/r3 === -Infinity ? "-0" : r3) + "," + (1/r4 === -Infinity ? "-0" : r4));
	`))
	if err != nil {
		t.Fatal(err)
	}
	expected := "-1,-2,-0,-0"
	if out != expected {
		t.Errorf("expected %q, got %q", expected, out)
	}
}

func TestJScriptMathConstants(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		Response.Write((Math.PI > 3 && Math.PI < 4 && Math.E > 2 && Math.E < 3) ? "true" : "false");
	`))
	if err != nil {
		t.Fatal(err)
	}
	expected := "true"
	if out != expected {
		t.Errorf("expected %q, got %q", expected, out)
	}
}

func TestJScriptMathMethods(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		Response.Write((Math.sin(0) === 0 && Math.cos(0) === 1 && Math.tan(0) === 0) ? "true" : "false");
	`))
	if err != nil {
		t.Fatal(err)
	}
	expected := "true"
	if out != expected {
		t.Errorf("expected %q, got %q", expected, out)
	}
}

// Phase 4: Array + Array concatenation

func TestJScriptArrayAddition(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		var result = [1, 2] + [3, 4];
		Response.Write(result);
	`))
	if err != nil {
		t.Fatal(err)
	}
	expected := "1,23,4"
	if out != expected {
		t.Errorf("expected %q, got %q", expected, out)
	}
}

// Phase 5: Eval, Function, wrapper equality, array equality

func TestJScriptArrayEqualsFalse(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		Response.Write(([] == false) ? "true" : "false");
	`))
	if err != nil {
		t.Fatal(err)
	}
	expected := "true"
	if out != expected {
		t.Errorf("expected %q, got %q", expected, out)
	}
}

func TestJScriptArrayEqualsZero(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		Response.Write(([] == 0) ? "true" : "false");
	`))
	if err != nil {
		t.Fatal(err)
	}
	expected := "true"
	if out != expected {
		t.Errorf("expected %q, got %q", expected, out)
	}
}

func TestJScriptArrayEqualsEmptyString(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		Response.Write(([] == "") ? "true" : "false");
	`))
	if err != nil {
		t.Fatal(err)
	}
	expected := "true"
	if out != expected {
		t.Errorf("expected %q, got %q", expected, out)
	}
}

func TestJScriptInstanceOfWrapperString(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		Response.Write((new String("x") instanceof String) ? "true" : "false");
	`))
	if err != nil {
		t.Fatal(err)
	}
	expected := "true"
	if out != expected {
		t.Errorf("expected %q, got %q", expected, out)
	}
}

func TestJScriptStringCoercionEmptyArray(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		Response.Write(String([]) + "|" + String([1, 2]));
	`))
	if err != nil {
		t.Fatal(err)
	}
	expected := "|1,2"
	if out != expected {
		t.Errorf("expected %q, got %q", expected, out)
	}
}

// Debug: exactly the failing expressions from the integration test

func TestJScriptNeg1Div0EqualsNegInfinity(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		Response.Write((-1 / 0 === -Infinity) ? "true" : "false");
	`))
	if err != nil {
		t.Fatal(err)
	}
	expected := "true"
	if out != expected {
		t.Errorf("expected %q, got %q", expected, out)
	}
}

func TestJScriptNewArray3Length(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		Response.Write((new Array(3)).length);
	`))
	if err != nil {
		t.Fatal(err)
	}
	expected := "3"
	if out != expected {
		t.Errorf("expected %q, got %q", expected, out)
	}
}

func TestJScriptAtan2PiOver2(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		Response.Write((Math.atan2(1, 0) === Math.PI / 2) ? "true" : "false");
	`))
	if err != nil {
		t.Fatal(err)
	}
	expected := "true"
	if out != expected {
		t.Errorf("expected %q, got %q", expected, out)
	}
}

func TestJScriptAtan2NegPiOver2(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		var v = Math.atan2(-1, 0);
		var pi2 = -Math.PI / 2;
		Response.Write("v=" + v + "|pi2=" + pi2 + "|eq=" + (v === pi2));
	`))
	if err != nil {
		t.Fatal(err)
	}
	expected := "v=-1.5707963267948966|pi2=-1.5707963267948966|eq=true"
	if out != expected {
		t.Errorf("expected %q, got %q", expected, out)
	}
}

func TestJScriptMathExpLog(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		Response.Write(Math.exp(0) + "|" + Math.log(1) + "|" + Math.log(0));
	`))
	if err != nil {
		t.Fatal(err)
	}
	expected := "1|0|-Infinity"
	if out != expected {
		t.Errorf("expected %q, got %q", expected, out)
	}
}

func TestJScriptMathExp(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		Response.Write(Math.exp(0));
	`))
	if err != nil {
		t.Fatal(err)
	}
	expected := "1"
	if out != expected {
		t.Errorf("expected %q, got %q", expected, out)
	}
}

func TestJScriptMathLog1(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		Response.Write(Math.log(1));
	`))
	if err != nil {
		t.Fatal(err)
	}
	expected := "0"
	if out != expected {
		t.Errorf("expected %q, got %q", expected, out)
	}
}

func TestJScriptMathLog0(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		Response.Write(Math.log(0));
	`))
	if err != nil {
		t.Fatal(err)
	}
	expected := "-Infinity"
	if out != expected {
		t.Errorf("expected %q, got %q", expected, out)
	}
}

func TestJScriptRoundNeg05Div(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		var r = Math.round(-0.5);
		var d = 1 / r;
		Response.Write("r=" + r + "|d=" + d + "|eq=" + (d === -Infinity));
	`))
	if err != nil {
		t.Fatal(err)
	}
	expected := "r=-0|d=-Infinity|eq=true"
	if out != expected {
		t.Errorf("expected %q, got %q", expected, out)
	}
}

func TestJScriptRoundNeg01Div(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		var r = Math.round(-0.1);
		var d = 1 / r;
		Response.Write("r=" + r + "|d=" + d + "|eq=" + (d === -Infinity));
	`))
	if err != nil {
		t.Fatal(err)
	}
	expected := "r=-0|d=-Infinity|eq=true"
	if out != expected {
		t.Errorf("expected %q, got %q", expected, out)
	}
}

func TestJScriptWrapperToStringTag(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		Response.Write(Object.prototype.toString.call(new String("x")) + "|" +
			Object.prototype.toString.call(new Number(1)) + "|" +
			Object.prototype.toString.call(new Boolean(false)));
	`))
	if err != nil {
		t.Fatal(err)
	}
	expected := "[object String]|[object Number]|[object Boolean]"
	if out != expected {
		t.Errorf("expected %q, got %q", expected, out)
	}
}

func TestJScriptArrayNotEqualsItsNegation(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		Response.Write(([] == ![]) ? "true" : "false");
	`))
	if err != nil {
		t.Fatal(err)
	}
	expected := "true"
	if out != expected {
		t.Errorf("expected %q, got %q", expected, out)
	}
}

func TestJScriptEval(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		var result = eval("1 + 2");
		Response.Write(result);
	`))
	if err != nil {
		t.Fatal(err)
	}
	expected := "3"
	if out != expected {
		t.Errorf("expected %q, got %q", expected, out)
	}
}

func TestJScriptStringCoercionArray(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		Response.Write(String([]) + "|" + String([1, 2]));
	`))
	if err != nil {
		t.Fatal(err)
	}
	expected := "|1,2"
	if out != expected {
		t.Errorf("expected %q, got %q", expected, out)
	}
}

func TestJScriptArrayAdditionWithObject(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		var r1 = [] + {};
		var r2 = {} + [];
		Response.Write(r1 + "|" + r2);
	`))
	if err != nil {
		t.Fatal(err)
	}
	expected := "[object Object]|[object Object]"
	if out != expected {
		t.Errorf("expected %q, got %q", expected, out)
	}
}

func TestJScriptMathAbs(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		var r1 = Math.abs("-1");
		var r2 = Math.abs("x");
		Response.Write((r1 === 1 && isNaN(r2) ? "true" : "false"));
	`))
	if err != nil {
		t.Fatal(err)
	}
	expected := "true"
	if out != expected {
		t.Errorf("expected %q, got %q", expected, out)
	}
}

// TestJScriptStringCoercionVTEmpty verifies that String(VTEmpty) returns "undefined"
// (VTEmpty maps to JS undefined, not null), per ES6+ String() semantics and IIS
// Classic ASP JScript bridge behavior.
func TestJScriptStringCoercionVTEmpty(t *testing.T) {
	vm := NewVM([]byte{}, nil, 0)
	host := NewMockHost()
	vm.SetHost(host)

	result := vm.jsToString(Value{Type: VTEmpty})
	expected := "undefined"
	if result != expected {
		t.Errorf("jsToString(VTEmpty) = %q, want %q", result, expected)
	}
}

// TestJScriptStringCoercionVTNull verifies no regression: String(null) still returns "null".
func TestJScriptStringCoercionVTNull(t *testing.T) {
	vm := NewVM([]byte{}, nil, 0)
	host := NewMockHost()
	vm.SetHost(host)

	result := vm.jsToString(Value{Type: VTNull})
	expected := "null"
	if result != expected {
		t.Errorf("jsToString(VTNull) = %q, want %q", result, expected)
	}
}

// TestJScriptStringCoercionVTJSUndefined verifies String(undefined) returns "undefined".
func TestJScriptStringCoercionVTJSUndefined(t *testing.T) {
	vm := NewVM([]byte{}, nil, 0)
	host := NewMockHost()
	vm.SetHost(host)

	result := vm.jsToString(Value{Type: VTJSUndefined})
	expected := "undefined"
	if result != expected {
		t.Errorf("jsToString(VTJSUndefined) = %q, want %q", result, expected)
	}
}

// TestJScriptNumberCoercionVTEmpty verifies that Number(VTEmpty) returns NaN
// (VTEmpty maps to JS undefined, and Number(undefined) = NaN in ES6+).
func TestJScriptNumberCoercionVTEmpty(t *testing.T) {
	vm := NewVM([]byte{}, nil, 0)
	host := NewMockHost()
	vm.SetHost(host)

	result := vm.jsToNumber(Value{Type: VTEmpty})
	if result.Type != VTDouble || !math.IsNaN(result.Flt) {
		t.Errorf("jsToNumber(VTEmpty) = %v, want NaN", result)
	}
}

// TestJScriptNumberCoercionVTNull verifies no regression: Number(null) still returns 0.
func TestJScriptNumberCoercionVTNull(t *testing.T) {
	vm := NewVM([]byte{}, nil, 0)
	host := NewMockHost()
	vm.SetHost(host)

	result := vm.jsToNumber(Value{Type: VTNull})
	if result.Type != VTDouble || result.Flt != 0 {
		t.Errorf("jsToNumber(VTNull) = %v, want 0", result)
	}
}

// TestJScriptNumberCoercionVTJSUndefined verifies Number(undefined) returns NaN.
func TestJScriptNumberCoercionVTJSUndefined(t *testing.T) {
	vm := NewVM([]byte{}, nil, 0)
	host := NewMockHost()
	vm.SetHost(host)

	result := vm.jsToNumber(Value{Type: VTJSUndefined})
	if result.Type != VTDouble || !math.IsNaN(result.Flt) {
		t.Errorf("jsToNumber(VTJSUndefined) = %v, want NaN", result)
	}
}

// TestJScriptMapKeyVTEmpty asserts VTEmpty and VTJSUndefined share the same Map key
// prefix ("u"), while VTNull gets its own distinct prefix ("n").
func TestJScriptMapKeyVTEmpty(t *testing.T) {
	vm := NewVM([]byte{}, nil, 0)
	host := NewMockHost()
	vm.SetHost(host)

	emptyKey := vm.jsValueMapKey(Value{Type: VTEmpty})
	undefKey := vm.jsValueMapKey(Value{Type: VTJSUndefined})
	nullKey := vm.jsValueMapKey(Value{Type: VTNull})

	if emptyKey != "u" {
		t.Errorf("jsValueMapKey(VTEmpty) = %q, want %q", emptyKey, "u")
	}
	if undefKey != "u" {
		t.Errorf("jsValueMapKey(VTJSUndefined) = %q, want %q", undefKey, "u")
	}
	if nullKey != "n" {
		t.Errorf("jsValueMapKey(VTNull) = %q, want %q", nullKey, "n")
	}
	if emptyKey != undefKey {
		t.Error("VTEmpty and VTJSUndefined must share the same Map key prefix")
	}
	if emptyKey == nullKey {
		t.Error("VTEmpty and VTNull must have distinct Map key prefixes")
	}
}

// TestJScriptStringCoercionViaASP verifies String() behavior via the ASP runtime
// for values that cross the VBScript/JScript bridge.
func TestJScriptStringCoercionViaASP(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		Response.Write("U:" + String(undefined) + "|N:" + String(null));
	`))
	if err != nil {
		t.Fatal(err)
	}
	expected := "U:undefined|N:null"
	if out != expected {
		t.Errorf("expected %q, got %q", expected, out)
	}
}

// TestJScriptRequestFormMissingField verifies that String(Request.Form("missing"))
// returns "" (empty string), matching IIS Classic ASP JScript behavior where
// missing form fields coerce to an empty string in string context.
func TestJScriptRequestFormMissingField(t *testing.T) {
	compiler := NewASPCompiler(jscriptSrc(`
		var v = Request.Form("no_such_field");
		Response.Write("T:" + typeof v + "|S:" + String(v));
	`))
	if err := compiler.Compile(); err != nil {
		t.Fatal(err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	host.Response().SetBuffer(false)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatal(err)
	}

	expected := "T:string|S:"
	if output.String() != expected {
		t.Errorf("expected %q, got %q", expected, output.String())
	}
}

// TestJScriptRequestFormEmptyField verifies that String(Request.Form("existing"))
// returns "" for a field submitted with an empty value.
func TestJScriptRequestFormEmptyField(t *testing.T) {
	compiler := NewASPCompiler(jscriptSrc(`
		Response.Write("[" + String(Request.Form("name")) + "]");
	`))
	if err := compiler.Compile(); err != nil {
		t.Fatal(err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	host.Response().SetBuffer(false)
	host.Request().Form.Add("name", "")
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatal(err)
	}

	expected := "[]"
	if output.String() != expected {
		t.Errorf("expected %q, got %q", expected, output.String())
	}
}

// TestJScriptRequestFormNonEmptyField verifies String(Request.Form("existing"))
// for a field with a real value returns the correct string.
func TestJScriptRequestFormNonEmptyField(t *testing.T) {
	compiler := NewASPCompiler(jscriptSrc(`
		Response.Write(String(Request.Form("email")));
	`))
	if err := compiler.Compile(); err != nil {
		t.Fatal(err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	host.Response().SetBuffer(false)
	host.Request().Form.Add("email", "user@test.com")
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatal(err)
	}

	expected := "user@test.com"
	if output.String() != expected {
		t.Errorf("expected %q, got %q", expected, output.String())
	}
}

// TestJScriptNumberCoercionViaASP verifies Number() behavior for undefined vs null
// via the ASP JScript runtime.
func TestJScriptNumberCoercionViaASP(t *testing.T) {
	out, err := runJScript2(t, jscriptSrc(`
		Response.Write("U:" + (isNaN(Number(undefined)) ? "NaN" : "NOT_NaN") +
			"|N:" + Number(null));
	`))
	if err != nil {
		t.Fatal(err)
	}
	expected := "U:NaN|N:0"
	if out != expected {
		t.Errorf("expected %q, got %q", expected, out)
	}
}

// TestValueToStringEmptyRequestCollection verifies that valueToString returns ""
// for a VTNativeObject wrapping an empty RequestCollectionValue, regardless of
// whether the caller is in JS or VBS mode.
func TestValueToStringEmptyRequestCollection(t *testing.T) {
	vm := NewVM([]byte{}, nil, 0)
	host := NewMockHost()
	vm.SetHost(host)

	// Simulate creating a request collection value item for a missing field.
	emptyVal := asp.RequestCollectionValue{} // len(Values) == 0
	id := vm.nextDynamicNativeID
	vm.nextDynamicNativeID++
	vm.requestCollectionValueItems[id] = emptyVal
	nativeVal := Value{Type: VTNativeObject, Num: id}

	result := vm.valueToString(nativeVal)
	if result != "" {
		t.Errorf("valueToString(empty RequestCollectionValue) = %q, want %q", result, "")
	}

	// Also verify jsToString path reaches the same result.
	jsResult := vm.jsToString(nativeVal)
	if jsResult != "" {
		t.Errorf("jsToString(empty RequestCollectionValue) = %q, want %q", jsResult, "")
	}
}
