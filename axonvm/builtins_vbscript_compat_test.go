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
	"strconv"
	"strings"
	"testing"
	"time"

	"g3pix.com.br/axonasp/v2/vbscript"
)

// callBuiltin resolves and executes one built-in by name for tests.
func callBuiltin(t *testing.T, vm *VM, name string, args ...Value) Value {
	t.Helper()
	idx, ok := GetBuiltinIndex(name)
	if !ok {
		t.Fatalf("builtin not found: %s", name)
	}
	result, err := BuiltinRegistry[idx](vm, args)
	if err != nil {
		t.Fatalf("builtin %s returned error: %v", name, err)
	}
	return result
}

// callBuiltinWithError resolves and executes one built-in by name for tests that expect runtime errors.
func callBuiltinWithError(t *testing.T, vm *VM, name string, args ...Value) (Value, error) {
	t.Helper()
	idx, ok := GetBuiltinIndex(name)
	if !ok {
		t.Fatalf("builtin not found: %s", name)
	}
	return BuiltinRegistry[idx](vm, args)
}

// TestBuiltinChrBMidBRoundTrip verifies byte-oriented string built-ins preserve raw byte values.
func TestBuiltinChrBMidBRoundTrip(t *testing.T) {
	vm := NewVM(nil, nil, 5)

	first := callBuiltin(t, vm, "ChrB", NewInteger(0x42))
	second := callBuiltin(t, vm, "ChrB", NewInteger(0x80))
	third := callBuiltin(t, vm, "ChrB", NewInteger(0xFF))
	combined := NewString(first.Str + second.Str + third.Str)

	mid := callBuiltin(t, vm, "MidB", combined, NewInteger(2), NewInteger(1))
	if mid.Type != VTString {
		t.Fatalf("expected VTString from MidB, got %#v", mid)
	}
	got := vbByteStringToBytes(mid.Str)
	if len(got) != 1 || got[0] != 0x80 {
		t.Fatalf("expected MidB byte 0x80, got %v", got)
	}

	asc := callBuiltin(t, vm, "AscB", mid)
	if asc.Type != VTInteger || asc.Num != 0x80 {
		t.Fatalf("expected AscB to round-trip 0x80, got %#v", asc)
	}
}

// TestBuiltinLocaleRoundTrip verifies SetLocale/GetLocale behavior against Session LCID.
func TestBuiltinLocaleRoundTrip(t *testing.T) {
	vm := NewVM(nil, nil, 5)
	host := NewMockHost()
	vm.SetHost(host)

	before := callBuiltin(t, vm, "GetLocale")
	if before.Type != VTInteger {
		t.Fatalf("expected integer GetLocale result, got %#v", before)
	}

	previous := callBuiltin(t, vm, "SetLocale", NewInteger(1046))
	if previous.Type != VTInteger || previous.Num != before.Num {
		t.Fatalf("unexpected SetLocale previous value: %#v", previous)
	}

	after := callBuiltin(t, vm, "GetLocale")
	if after.Type != VTInteger || after.Num != 1046 {
		t.Fatalf("expected updated LCID 1046, got %#v", after)
	}
}

// TestBuiltinDefaultLocaleUsesConfiguredLCID verifies built-ins pick up global.default_mslcid when no request override exists.
func TestBuiltinDefaultLocaleUsesConfiguredLCID(t *testing.T) {
	vm := NewVM(nil, nil, 5)
	host := NewMockHost()
	vm.SetHost(host)

	expected := int64(loadBuiltinDefaults().mslcid)
	configured := callBuiltin(t, vm, "GetLocale")
	if configured.Type != VTInteger || configured.Num != expected {
		t.Fatalf("expected configured default LCID %d, got %#v", expected, configured)
	}
}

// TestBuiltinFormatDateTimeUsesPortugueseLocale verifies vbLongDate and vbShortDate honor the configured Portuguese locale.
func TestBuiltinFormatDateTimeUsesPortugueseLocale(t *testing.T) {
	vm := NewVM(nil, nil, 5)
	host := NewMockHost()
	vm.SetHost(host)
	host.Session().SetLCID(int(PortugueseBrazil))

	dateVal := NewDate(time.Date(2026, time.April, 9, 16, 5, 7, 0, time.UTC))
	longDate := callBuiltin(t, vm, "FormatDateTime", dateVal, NewInteger(1))
	shortDate := callBuiltin(t, vm, "FormatDateTime", dateVal, NewInteger(2))
	longTime := callBuiltin(t, vm, "FormatDateTime", dateVal, NewInteger(3))

	if longDate.Type != VTString || !strings.Contains(strings.ToLower(longDate.Str), "abril") {
		t.Fatalf("expected Portuguese long date, got %#v", longDate)
	}
	if shortDate.Type != VTString || shortDate.Str != "09/04/2026" {
		t.Fatalf("expected Brazilian short date 09/04/2026, got %#v", shortDate)
	}
	expectedLongTime := valueToTimeInLocale(vm, dateVal).In(builtinCurrentLocation(vm)).Format("15:04:05")
	if longTime.Type != VTString || longTime.Str != expectedLongTime {
		t.Fatalf("expected 24-hour long time in server timezone %q, got %#v", expectedLongTime, longTime)
	}
}

// TestBuiltinFormatCurrencyUsesPortugueseLocale verifies currency formatting honors decimal and thousand separators from LCID.
func TestBuiltinFormatCurrencyUsesPortugueseLocale(t *testing.T) {
	vm := NewVM(nil, nil, 5)
	host := NewMockHost()
	vm.SetHost(host)
	host.Session().SetLCID(int(PortugueseBrazil))

	formatted := callBuiltin(t, vm, "FormatCurrency", NewDouble(1234.5))
	if formatted.Type != VTString || formatted.Str != "R$ 1.234,50" {
		t.Fatalf("expected Brazilian currency format, got %#v", formatted)
	}

	implicitDate := time.Date(2026, time.April, 9, 0, 0, 0, 0, builtinCurrentLocation(vm))
	implicit := vm.valueToString(NewDate(implicitDate))
	if implicit != "09/04/2026" {
		t.Fatalf("expected implicit localized date string 09/04/2026, got %q", implicit)
	}
}

// TestBuiltinCreateObjectCompatibility verifies built-in CreateObject delegates to Server.CreateObject.
func TestBuiltinCreateObjectCompatibility(t *testing.T) {
	vm := NewVM(nil, nil, 5)
	host := NewMockHost()
	vm.SetHost(host)

	obj := callBuiltin(t, vm, "CreateObject", NewString("G3Crypto"))
	if obj.Type != VTNativeObject {
		t.Fatalf("expected VTNativeObject from CreateObject, got %#v", obj)
	}
}

// TestBuiltinTypeNameDictionary verifies native dictionary TypeName compatibility.
func TestBuiltinTypeNameDictionary(t *testing.T) {
	vm := NewVM(nil, nil, 5)
	host := NewMockHost()
	vm.SetHost(host)

	dict := vm.newDictionaryObject()
	result := callBuiltin(t, vm, "TypeName", dict)
	if result.Type != VTString || result.Str != "Dictionary" {
		t.Fatalf("expected TypeName(Dictionary)='Dictionary', got %#v", result)
	}
}

// TestBuiltinBinaryInStrB verifies byte-oriented search behavior with 1-based index.
func TestBuiltinBinaryInStrB(t *testing.T) {
	vm := NewVM(nil, nil, 5)
	result := callBuiltin(t, vm, "InStrB", NewString("AxonASP"), NewString("ASP"))
	if result.Type != VTInteger || result.Num != 5 {
		t.Fatalf("expected InStrB result 5, got %#v", result)
	}
}

// TestBuiltinDateMath verifies DateAdd/DateDiff compatibility behavior.
func TestBuiltinDateMath(t *testing.T) {
	vm := NewVM(nil, nil, 5)

	base := NewString("2026-03-10 00:00:00")
	added := callBuiltin(t, vm, "DateAdd", NewString("d"), NewInteger(5), base)
	if added.Type != VTDate {
		t.Fatalf("expected VTDate from DateAdd, got %#v", added)
	}

	diff := callBuiltin(t, vm, "DateDiff", NewString("d"), base, added)
	if diff.Type != VTInteger || diff.Num != 5 {
		t.Fatalf("expected DateDiff=5, got %#v", diff)
	}
}

// TestBuiltinDateDiffCalendarDay verifies that DateDiff("d") counts crossed calendar midnights.
func TestBuiltinDateDiffCalendarDay(t *testing.T) {
	vm := NewVM(nil, nil, 5)

	start := NewString("2026-03-10 23:59:59")
	finish := NewString("2026-03-11 00:00:00")
	diff := callBuiltin(t, vm, "DateDiff", NewString("d"), start, finish)
	if diff.Type != VTInteger || diff.Num != 1 {
		t.Fatalf("expected DateDiff(\"d\")=1 across midnight, got %#v", diff)
	}
}

// TestBuiltinDateSerialCoercesStringArgs verifies Classic VBScript behavior where
// DateSerial accepts numeric strings from form/query data.
func TestBuiltinDateSerialCoercesStringArgs(t *testing.T) {
	vm := NewVM(nil, nil, 5)

	dateVal := callBuiltin(t, vm, "DateSerial", NewString("1974"), NewString("07"), NewString("10"))
	if dateVal.Type != VTDate {
		t.Fatalf("expected VTDate from DateSerial string args, got %#v", dateVal)
	}

	formatted := callBuiltin(t, vm, "FormatDateTime", dateVal, NewInteger(2))
	if formatted.Type != VTString {
		t.Fatalf("expected VTString from FormatDateTime, got %#v", formatted)
	}
	if !strings.Contains(formatted.Str, "1974") {
		t.Fatalf("expected formatted date to include 1974, got %q", formatted.Str)
	}
}

// TestBuiltinIsDateRejectsZeroDateSentinel verifies the internal zero-date
// sentinel is not treated as a valid VBScript date.
func TestBuiltinIsDateRejectsZeroDateSentinel(t *testing.T) {
	vm := NewVM(nil, nil, 5)

	zeroDate := NewDate(time.Time{})
	result := callBuiltin(t, vm, "IsDate", zeroDate)
	if result.Type != VTBool || result.Num != 0 {
		t.Fatalf("expected IsDate(NewDate(time.Time{}))=False, got %#v", result)
	}
}

// TestBuiltinInStrRuneStart verifies InStr start handling with multi-byte characters.
func TestBuiltinInStrRuneStart(t *testing.T) {
	vm := NewVM(nil, nil, 5)

	result := callBuiltin(t, vm, "InStr", NewInteger(2), NewString("ááb"), NewString("b"))
	if result.Type != VTInteger || result.Num != 3 {
		t.Fatalf("expected InStr rune-safe result 3, got %#v", result)
	}
}

// TestBuiltinInStrTextCompare verifies explicit text compare behavior for ASCII and Unicode.
func TestBuiltinInStrTextCompare(t *testing.T) {
	vm := NewVM(nil, nil, 5)

	ascii := callBuiltin(t, vm, "InStr", NewInteger(1), NewString("ABC"), NewString("a"), NewInteger(1))
	if ascii.Type != VTInteger || ascii.Num != 1 {
		t.Fatalf("expected InStr text compare ASCII result 1, got %#v", ascii)
	}

	unicode := callBuiltin(t, vm, "InStr", NewInteger(1), NewString("Árvore"), NewString("á"), NewInteger(1))
	if unicode.Type != VTInteger || unicode.Num != 1 {
		t.Fatalf("expected InStr text compare Unicode result 1, got %#v", unicode)
	}
}

// TestBuiltinInStrInvalidStartError verifies that start <= 0 raises VBScript invalid-procedure error.
func TestBuiltinInStrInvalidStartError(t *testing.T) {
	vm := NewVM(nil, nil, 5)

	_, err := callBuiltinWithError(t, vm, "InStr", NewInteger(0), NewString("abc"), NewString("a"))
	if err == nil {
		t.Fatalf("expected runtime error for InStr start=0")
	}
	runtimeErr, ok := err.(builtinVBRuntimeError)
	if !ok {
		t.Fatalf("expected builtinVBRuntimeError, got %T", err)
	}
	if runtimeErr.code != vbscript.InvalidProcedureCallOrArgument {
		t.Fatalf("expected InvalidProcedureCallOrArgument, got %v", runtimeErr.code)
	}
}

// TestBuiltinReplaceStartSemantics verifies VBScript Replace result starts at the start position.
func TestBuiltinReplaceStartSemantics(t *testing.T) {
	vm := NewVM(nil, nil, 5)

	result := callBuiltin(t, vm, "Replace", NewString("abcabc"), NewString("a"), NewString("X"), NewInteger(2))
	if result.Type != VTString || result.Str != "bcXbc" {
		t.Fatalf("expected Replace start-semantics output bcXbc, got %#v", result)
	}
}

// TestBuiltinReplaceNullExpression verifies Replace raises 94 for a Null
// expression. Measured on IIS: Replace is the exception among the string
// functions, which return Null for a Null argument
// (https://github.com/guimaraeslucas/axonasp/issues/188).
func TestBuiltinReplaceNullExpression(t *testing.T) {
	vm := NewVM(nil, nil, 5)

	_, err := callBuiltinWithError(t, vm, "Replace", NewNull(), NewString("a"), NewString("b"))
	if err == nil || !strings.Contains(err.Error(), "Invalid use of Null") {
		t.Fatalf("expected Invalid use of Null from Replace(Null,...), got %v", err)
	}
}

// TestBuiltinReplaceInvalidStartError verifies that Replace start <= 0 raises VBScript invalid-procedure error.
func TestBuiltinReplaceInvalidStartError(t *testing.T) {
	vm := NewVM(nil, nil, 5)

	_, err := callBuiltinWithError(t, vm, "Replace", NewString("abc"), NewString("a"), NewString("b"), NewInteger(0))
	if err == nil {
		t.Fatalf("expected runtime error for Replace start=0")
	}
	runtimeErr, ok := err.(builtinVBRuntimeError)
	if !ok {
		t.Fatalf("expected builtinVBRuntimeError, got %T", err)
	}
	if runtimeErr.code != vbscript.InvalidProcedureCallOrArgument {
		t.Fatalf("expected InvalidProcedureCallOrArgument, got %v", runtimeErr.code)
	}
}

// TestBuiltinReplaceCountLimit verifies replacement stops after the requested count.
func TestBuiltinReplaceCountLimit(t *testing.T) {
	vm := NewVM(nil, nil, 5)

	result := callBuiltin(t, vm, "Replace", NewString("aaaa"), NewString("a"), NewString("x"), NewInteger(1), NewInteger(2), NewInteger(0))
	if result.Type != VTString || result.Str != "xxaa" {
		t.Fatalf("expected count-limited output xxaa, got %#v", result)
	}
}

// TestBuiltinReplaceBinaryCompare verifies case-sensitive behavior for compare=0.
func TestBuiltinReplaceBinaryCompare(t *testing.T) {
	vm := NewVM(nil, nil, 5)

	result := callBuiltin(t, vm, "Replace", NewString("aAaA"), NewString("a"), NewString("x"), NewInteger(1), NewInteger(-1), NewInteger(0))
	if result.Type != VTString || result.Str != "xAxA" {
		t.Fatalf("expected binary-compare output xAxA, got %#v", result)
	}
}

// TestBuiltinReplaceTextCompare verifies case-insensitive behavior for compare=1.
func TestBuiltinReplaceTextCompare(t *testing.T) {
	vm := NewVM(nil, nil, 5)

	result := callBuiltin(t, vm, "Replace", NewString("aAaA"), NewString("a"), NewString("x"), NewInteger(1), NewInteger(-1), NewInteger(1))
	if result.Type != VTString || result.Str != "xxxx" {
		t.Fatalf("expected text-compare output xxxx, got %#v", result)
	}
}

// TestBuiltinReplaceTextCompareUnicode verifies Unicode-aware text compare behavior.
func TestBuiltinReplaceTextCompareUnicode(t *testing.T) {
	vm := NewVM(nil, nil, 5)

	result := callBuiltin(t, vm, "Replace", NewString("ÁáÀ"), NewString("á"), NewString("x"), NewInteger(1), NewInteger(-1), NewInteger(1))
	if result.Type != VTString || result.Str != "xxÀ" {
		t.Fatalf("expected Unicode text-compare output xxÀ, got %#v", result)
	}
}

// TestBuiltinReplaceEmptyFindReturnsTail verifies empty find keeps the substring from start unchanged.
func TestBuiltinReplaceEmptyFindReturnsTail(t *testing.T) {
	vm := NewVM(nil, nil, 5)

	result := callBuiltin(t, vm, "Replace", NewString("abc"), NewString(""), NewString("X"), NewInteger(2), NewInteger(-1), NewInteger(0))
	if result.Type != VTString || result.Str != "bc" {
		t.Fatalf("expected empty-find output bc, got %#v", result)
	}
}

// TestBuiltinInteractiveDesktopFunctionsRejectASP verifies MsgBox/InputBox fail in server-side ASP.
func TestBuiltinInteractiveDesktopFunctionsRejectASP(t *testing.T) {
	vm := NewVM(nil, nil, 5)

	tests := []string{"MsgBox", "InputBox"}
	for _, name := range tests {
		_, err := callBuiltinWithError(t, vm, name, NewString("prompt"))
		if err == nil {
			t.Fatalf("expected runtime error for %s in ASP", name)
		}
		runtimeErr, ok := err.(builtinVBRuntimeError)
		if !ok {
			t.Fatalf("expected builtinVBRuntimeError for %s, got %T", name, err)
		}
		if runtimeErr.code != vbscript.InvalidProcedureCallOrArgument {
			t.Fatalf("expected InvalidProcedureCallOrArgument for %s, got %v", name, runtimeErr.code)
		}
		if !strings.Contains(runtimeErr.Error(), ErrInteractiveFunctionNotSupportedInASP.String()) {
			t.Fatalf("expected AxonASP interactive-function message for %s, got %q", name, runtimeErr.Error())
		}
	}
}

// TestBuiltinStrConvWidthModes verifies best-effort width conversions for StrConv.
func TestBuiltinStrConvWidthModes(t *testing.T) {
	vm := NewVM(nil, nil, 5)

	wide := callBuiltin(t, vm, "StrConv", NewString("ABC 123"), NewInteger(4))
	if wide.Type != VTString || wide.Str != "ＡＢＣ　１２３" {
		t.Fatalf("expected vbWide conversion, got %#v", wide)
	}

	narrow := callBuiltin(t, vm, "StrConv", NewString("ＡＢＣ　１２３"), NewInteger(8))
	if narrow.Type != VTString || narrow.Str != "ABC 123" {
		t.Fatalf("expected vbNarrow conversion, got %#v", narrow)
	}

	fromUnicode := callBuiltin(t, vm, "StrConv", NewString("ＡＢＣ　１２３"), NewInteger(128))
	if fromUnicode.Type != VTString || fromUnicode.Str != "ABC 123" {
		t.Fatalf("expected vbFromUnicode best-effort conversion, got %#v", fromUnicode)
	}
}

// TestBuiltinSplitJoinFilter verifies array string helper compatibility.
func TestBuiltinSplitJoinFilter(t *testing.T) {
	vm := NewVM(nil, nil, 5)
	split := callBuiltin(t, vm, "Split", NewString("a,b,c"), NewString(","))
	if split.Type != VTArray || split.Arr == nil || len(split.Arr.Values) != 3 {
		t.Fatalf("unexpected Split output: %#v", split)
	}
	joined := callBuiltin(t, vm, "Join", split, NewString("|"))
	if joined.Type != VTString || joined.Str != "a|b|c" {
		t.Fatalf("unexpected Join output: %#v", joined)
	}
	filtered := callBuiltin(t, vm, "Filter", split, NewString("b"), NewBool(true), NewInteger(0))
	if filtered.Type != VTArray || filtered.Arr == nil || len(filtered.Arr.Values) != 1 || filtered.Arr.Values[0].Str != "b" {
		t.Fatalf("unexpected Filter output: %#v", filtered)
	}
}

// TestBuiltinRoundBankers verifies VBScript/IIS banker's rounding behavior.
func TestBuiltinRoundBankers(t *testing.T) {
	vm := NewVM(nil, nil, 5)

	r1 := callBuiltin(t, vm, "Round", NewDouble(6.5))
	if r1.Type != VTDouble || r1.Flt != 6 {
		t.Fatalf("expected Round(6.5)=6, got %#v", r1)
	}

	r2 := callBuiltin(t, vm, "Round", NewDouble(7.5))
	if r2.Type != VTDouble || r2.Flt != 8 {
		t.Fatalf("expected Round(7.5)=8, got %#v", r2)
	}

	r3 := callBuiltin(t, vm, "Round", NewDouble(2.25), NewInteger(1))
	if r3.Type != VTDouble || r3.Flt != 2.2 {
		t.Fatalf("expected Round(2.25, 1)=2.2, got %#v", r3)
	}

	r4 := callBuiltin(t, vm, "Round", NewDouble(2.35), NewInteger(1))
	if r4.Type != VTDouble || r4.Flt != 2.4 {
		t.Fatalf("expected Round(2.35, 1)=2.4, got %#v", r4)
	}
}

// TestBuiltinHexOctDoubleCompat verifies that Hex/Oct accept VTDouble (result of VBScript /
// division) without returning zero — the root cause of BMP header corruption in SendHex.
func TestBuiltinHexOctDoubleCompat(t *testing.T) {
	vm := NewVM(nil, nil, 5)

	// Hex with VTInteger — baseline
	h1 := callBuiltin(t, vm, "Hex", NewInteger(255))
	if h1.Type != VTString || h1.Str != "FF" {
		t.Fatalf("expected Hex(255)=\"FF\", got %#v", h1)
	}

	// Hex with VTDouble (VBScript / always returns Double) — was returning "0" before fix
	h2 := callBuiltin(t, vm, "Hex", NewDouble(18.0))
	if h2.Type != VTString || h2.Str != "12" {
		t.Fatalf("expected Hex(18.0)=\"12\", got %#v", h2)
	}

	// Hex with VTDouble that needs rounding
	h3 := callBuiltin(t, vm, "Hex", NewDouble(255.5)) // rounds to 256 = 0x100
	if h3.Type != VTString || h3.Str != "100" {
		t.Fatalf("expected Hex(255.5)=\"100\", got %#v", h3)
	}

	// BMP header use-case: Len("...") / 2 produces Double 14.0
	h4 := callBuiltin(t, vm, "Hex", NewDouble(126.0))
	if h4.Type != VTString || h4.Str != "7E" {
		t.Fatalf("expected Hex(126.0)=\"7E\", got %#v", h4)
	}

	// Oct with VTDouble
	o1 := callBuiltin(t, vm, "Oct", NewDouble(8.0))
	if o1.Type != VTString || o1.Str != "10" {
		t.Fatalf("expected Oct(8.0)=\"10\", got %#v", o1)
	}
}

// TestBuiltinFormatNumberUS verifies US English number formatting.
func TestBuiltinFormatNumberUS(t *testing.T) {
	vm := NewVM(nil, nil, 5)
	host := NewMockHost()
	vm.SetHost(host)
	host.Session().SetLCID(int(EnglishUS))

	// US: decimal separator ".", thousands separator ","
	result := callBuiltin(t, vm, "FormatNumber", NewDouble(1234.567), NewInteger(2))
	if result.Type != VTString || result.Str != "1,234.57" {
		t.Fatalf("expected US FormatNumber(1234.567,2)='1,234.57', got %q", result.Str)
	}
}

// TestBuiltinFormatNumberDE verifies German number formatting.
func TestBuiltinFormatNumberDE(t *testing.T) {
	vm := NewVM(nil, nil, 5)
	host := NewMockHost()
	vm.SetHost(host)
	host.Session().SetLCID(int(GermanGermany))

	// DE: decimal separator ",", thousands separator "."
	result := callBuiltin(t, vm, "FormatNumber", NewDouble(1234.567), NewInteger(2))
	if result.Type != VTString || result.Str != "1.234,57" {
		t.Fatalf("expected German FormatNumber(1234.567,2)='1.234,57', got %q", result.Str)
	}
}

// TestBuiltinFormatNumberFR verifies French number formatting with space separator.
func TestBuiltinFormatNumberFR(t *testing.T) {
	vm := NewVM(nil, nil, 5)
	host := NewMockHost()
	vm.SetHost(host)
	host.Session().SetLCID(int(FrenchFrance))

	// FR: decimal separator ",", thousands separator " " (space)
	result := callBuiltin(t, vm, "FormatNumber", NewDouble(1234567.5), NewInteger(1))
	if result.Type != VTString || result.Str != "1 234 567,5" {
		t.Fatalf("expected French FormatNumber(1234567.5,1)='1 234 567,5', got %q", result.Str)
	}
}

// TestBuiltinFormatPercentDE verifies German percentage formatting.
func TestBuiltinFormatPercentDE(t *testing.T) {
	vm := NewVM(nil, nil, 5)
	host := NewMockHost()
	vm.SetHost(host)
	host.Session().SetLCID(int(GermanGermany))

	// DE: FormatPercent multiplies by 100, appends "%", uses German separators
	result := callBuiltin(t, vm, "FormatPercent", NewDouble(0.1234), NewInteger(1))
	if result.Type != VTString || result.Str != "12,3%" {
		t.Fatalf("expected German FormatPercent(0.1234,1)='12,3%%', got %q", result.Str)
	}
}

// TestBuiltinFormatPercentFR verifies French percentage formatting.
func TestBuiltinFormatPercentFR(t *testing.T) {
	vm := NewVM(nil, nil, 5)
	host := NewMockHost()
	vm.SetHost(host)
	host.Session().SetLCID(int(FrenchFrance))

	result := callBuiltin(t, vm, "FormatPercent", NewDouble(0.5), NewInteger(0))
	if result.Type != VTString || result.Str != "50%" {
		t.Fatalf("expected French FormatPercent(0.5,0)='50%%', got %q", result.Str)
	}
}

func assertFormatBuiltinRuntimeErr(t *testing.T, vm *VM, fnName string, expectedCode vbscript.VBSyntaxErrorCode, args ...Value) {
	t.Helper()
	_, err := callBuiltinWithError(t, vm, fnName, args...)
	if err == nil {
		t.Fatalf("%s(%v): expected error %v, got nil", fnName, args, expectedCode)
	}
	vbErr, ok := err.(builtinVBRuntimeError)
	if !ok {
		t.Fatalf("%s(%v): expected builtinVBRuntimeError, got %T (%v)", fnName, args, err, err)
	}
	if vbErr.code != expectedCode {
		t.Fatalf("%s(%v): expected error code %d (%v), got %d (%v)", fnName, args, expectedCode, expectedCode, vbErr.code, vbErr.code)
	}
}

func TestBuiltinFormatTypeHandlingAndParsing(t *testing.T) {
	host := NewMockHost()
	host.Session().SetLCID(1033)
	vm := NewVM(nil, nil, 0)
	vm.SetHost(host)

	// VTInteger, VTDouble, VTEmpty, VTBool, and trimmed/scientific/grouped VTString inputs.
	cases := []struct {
		fn       string
		args     []Value
		expected string
	}{
		{"FormatCurrency", []Value{NewInteger(1234)}, "$1,234.00"},
		{"FormatCurrency", []Value{NewDouble(1234.5)}, "$1,234.50"},
		{"FormatCurrency", []Value{NewEmpty()}, "$0.00"},
		{"FormatCurrency", []Value{NewString("1234.50")}, "$1,234.50"},
		{"FormatCurrency", []Value{NewString("  1,234.50  ")}, "$1,234.50"},
		{"FormatCurrency", []Value{NewString("-1234.5")}, "-$1,234.50"},
		{"FormatCurrency", []Value{NewString("+42.5")}, "$42.50"},
		{"FormatCurrency", []Value{NewString(".5")}, "$0.50"},
		{"FormatCurrency", []Value{NewString("1.2345e2")}, "$123.45"},
		{"FormatCurrency", []Value{NewString("1.2345D2")}, "$123.45"},
		{"FormatCurrency", []Value{NewString("1.5d-1")}, "$0.15"},
		{"FormatCurrency", []Value{NewString("&H10")}, "$16.00"},
		{"FormatCurrency", []Value{NewString("&O10")}, "$8.00"},
		{"FormatNumber", []Value{NewString("1,234.5")}, "1,234.50"},
		{"FormatNumber", []Value{NewEmpty()}, "0.00"},
		{"FormatNumber", []Value{NewBool(false)}, "0.00"},
		{"FormatNumber", []Value{NewBool(true)}, "1.00"},
		{"FormatNumber", []Value{NewString("1e10"), NewInteger(0)}, "10,000,000,000"},
		{"FormatNumber", []Value{NewString("1e-5"), NewInteger(5)}, "0.00001"},
		{"FormatNumber", []Value{NewString("1e20"), NewInteger(0)}, "100,000,000,000,000,000,000"},
		{"FormatPercent", []Value{NewString("0.125")}, "12.50%"},
		{"FormatPercent", []Value{NewString("1.25"), NewInteger(0)}, "125%"},
		{"FormatPercent", []Value{NewEmpty()}, "0.00%"},
		{"FormatPercent", []Value{NewString("0.5")}, "50.00%"},
	}
	for _, tc := range cases {
		res := callBuiltin(t, vm, tc.fn, tc.args...)
		if res.Type != VTString || res.Str != tc.expected {
			t.Fatalf("%s(%v): expected %q, got Type=%v Str=%q", tc.fn, tc.args, tc.expected, res.Type, res.Str)
		}
	}

	// VTNull returns Null for FormatCurrency, FormatNumber, and FormatPercent.
	for _, fn := range []string{"FormatCurrency", "FormatNumber", "FormatPercent"} {
		res := callBuiltin(t, vm, fn, NewNull())
		if res.Type != VTNull {
			t.Fatalf("%s(Null): expected VTNull, got %v (%q)", fn, res.Type, res.Str)
		}
	}

	// Invalid strings and non-numeric types raise vbscript.TypeMismatch.
	invalidArgs := []Value{
		NewString(""),
		NewString("   "),
		NewString("abc"),
		NewString("123abc"),
		NewString("12.34.56"),
		NewString("NaN"),
		NewString("Inf"),
		NewString("Infinity"),
		NewString("1_000"),
		NewString("0x10"),
		NewString("0x1p2"),
		NewString("+-1"),
		NewString("12,34"),
		NewString("&H"),
		NewString("&O9"),
		NewDouble(math.NaN()),
		{Type: VTArray},
		{Type: VTNothing},
	}
	for _, bad := range invalidArgs {
		for _, fn := range []string{"FormatCurrency", "FormatNumber", "FormatPercent"} {
			assertFormatBuiltinRuntimeErr(t, vm, fn, vbscript.TypeMismatch, bad)
		}
	}

	// Infinity double or huge numeric string raises vbscript.Overflow.
	for _, fn := range []string{"FormatCurrency", "FormatNumber", "FormatPercent"} {
		assertFormatBuiltinRuntimeErr(t, vm, fn, vbscript.Overflow, NewDouble(math.Inf(1)))
		assertFormatBuiltinRuntimeErr(t, vm, fn, vbscript.Overflow, NewDouble(math.Inf(-1)))
		assertFormatBuiltinRuntimeErr(t, vm, fn, vbscript.Overflow, NewString("1e400"))
	}
}

func TestBuiltinFormatNegativeZeroAndBankersRounding(t *testing.T) {
	host := NewMockHost()
	host.Session().SetLCID(1033)
	vm := NewVM(nil, nil, 0)
	vm.SetHost(host)

	cases := []struct {
		fn       string
		args     []Value
		expected string
	}{
		// Negative zero suppression (direct negative zero and values rounding to zero).
		{"FormatNumber", []Value{NewDouble(math.Copysign(0, -1))}, "0.00"},
		{"FormatNumber", []Value{NewString("-0")}, "0.00"},
		{"FormatNumber", []Value{NewString("-0.00")}, "0.00"},
		{"FormatNumber", []Value{NewDouble(-0.001), NewInteger(2)}, "0.00"},
		{"FormatNumber", []Value{NewDouble(-0.4), NewInteger(0)}, "0"},
		{"FormatCurrency", []Value{NewDouble(-0.001), NewInteger(2), NewInteger(-2), NewInteger(-1)}, "$0.00"},
		{"FormatPercent", []Value{NewDouble(-0.00001), NewInteger(2), NewInteger(-2), NewInteger(-1)}, "0.00%"},

		// Microsoft VBScript banker's rounding (round half to even).
		{"FormatNumber", []Value{NewDouble(1.225), NewInteger(2)}, "1.22"},
		{"FormatNumber", []Value{NewDouble(1.235), NewInteger(2)}, "1.24"},
		{"FormatNumber", []Value{NewDouble(2.675), NewInteger(2)}, "2.68"},
		{"FormatNumber", []Value{NewDouble(6.5), NewInteger(0)}, "6"},
		{"FormatNumber", []Value{NewDouble(7.5), NewInteger(0)}, "8"},
		{"FormatNumber", []Value{NewDouble(0.5), NewInteger(0)}, "0"},
		{"FormatNumber", []Value{NewDouble(0.005), NewInteger(2)}, "0.00"},
		{"FormatNumber", []Value{NewDouble(9.995), NewInteger(2)}, "10.00"},
		{"FormatPercent", []Value{NewString("0.12345"), NewInteger(2)}, "12.34%"},
		{"FormatPercent", []Value{NewString("0.12355"), NewInteger(2)}, "12.36%"},
	}
	for _, tc := range cases {
		res := callBuiltin(t, vm, tc.fn, tc.args...)
		if res.Type != VTString || res.Str != tc.expected {
			t.Fatalf("%s(%v): expected %q, got %q", tc.fn, tc.args, tc.expected, res.Str)
		}
	}
}

func TestBuiltinFormatOptionalArgumentsAndLocales(t *testing.T) {
	host := NewMockHost()
	vm := NewVM(nil, nil, 0)
	vm.SetHost(host)

	// All 5 arguments in en-US (LCID 1033).
	host.Session().SetLCID(1033)
	optCases := []struct {
		fn       string
		args     []Value
		expected string
	}{
		// NumDigitsAfterDecimal = -1 uses default (2).
		{"FormatNumber", []Value{NewString("12.3"), NewInteger(-1)}, "12.30"},
		// IncludeLeadingDigit = vbFalse (0) vs vbTrue (-1).
		{"FormatNumber", []Value{NewString("0.5"), NewInteger(2), NewInteger(0)}, ".50"},
		{"FormatNumber", []Value{NewString("-0.5"), NewInteger(2), NewInteger(0)}, "-.50"},
		{"FormatNumber", []Value{NewString("-0.5"), NewInteger(2), NewInteger(0), NewInteger(-1)}, "(.50)"},
		{"FormatCurrency", []Value{NewString("0.5"), NewInteger(2), NewInteger(0)}, "$.50"},
		{"FormatPercent", []Value{NewString("0.005"), NewInteger(2), NewInteger(0)}, ".50%"},
		// UseParensForNegativeNumbers = vbTrue (-1) vs vbFalse (0) vs vbUseDefault (-2).
		{"FormatCurrency", []Value{NewString("-1234.5"), NewInteger(2), NewInteger(-2), NewInteger(-1)}, "($1,234.50)"},
		{"FormatCurrency", []Value{NewString("-1234.5"), NewInteger(2), NewInteger(-2), NewInteger(0)}, "-$1,234.50"},
		{"FormatCurrency", []Value{NewString("-1234.5"), NewInteger(2), NewInteger(-2), NewInteger(-2)}, "-$1,234.50"},
		{"FormatNumber", []Value{NewString("-1234.5"), NewInteger(2), NewInteger(-2), NewInteger(-1)}, "(1,234.50)"},
		{"FormatPercent", []Value{NewString("-0.125"), NewInteger(1), NewInteger(-2), NewInteger(-1)}, "(12.5%)"},
		// GroupDigits = vbFalse (0) vs vbTrue (-1).
		{"FormatNumber", []Value{NewString("1234.5"), NewInteger(2), NewInteger(-2), NewInteger(-2), NewInteger(0)}, "1234.50"},
		{"FormatCurrency", []Value{NewString("1234.5"), NewInteger(2), NewInteger(-2), NewInteger(-2), NewInteger(0)}, "$1234.50"},
		{"FormatPercent", []Value{NewString("12.345"), NewInteger(1), NewInteger(-2), NewInteger(-2), NewInteger(0)}, "1234.5%"},
		{"FormatPercent", []Value{NewString("12.345"), NewInteger(1), NewInteger(-2), NewInteger(-2), NewInteger(-1)}, "1,234.5%"},
	}
	for _, tc := range optCases {
		res := callBuiltin(t, vm, tc.fn, tc.args...)
		if res.Type != VTString || res.Str != tc.expected {
			t.Fatalf("%s(%v): expected %q, got %q", tc.fn, tc.args, tc.expected, res.Str)
		}
	}

	// Optional parameter error conditions.
	assertFormatBuiltinRuntimeErr(t, vm, "FormatNumber", vbscript.InvalidProcedureCallOrArgument, NewInteger(1), NewInteger(-2))
	assertFormatBuiltinRuntimeErr(t, vm, "FormatNumber", vbscript.InvalidProcedureCallOrArgument, NewInteger(1), NewInteger(100))
	assertFormatBuiltinRuntimeErr(t, vm, "FormatNumber", vbscript.InvalidUseOfNull, NewInteger(1), NewNull())
	assertFormatBuiltinRuntimeErr(t, vm, "FormatCurrency", vbscript.InvalidUseOfNull, NewInteger(1), NewInteger(2), NewNull())
	assertFormatBuiltinRuntimeErr(t, vm, "FormatPercent", vbscript.InvalidProcedureCallOrArgument, NewInteger(1), NewInteger(2), NewInteger(5))
	assertFormatBuiltinRuntimeErr(t, vm, "FormatNumber", vbscript.InvalidProcedureCallOrArgument, NewInteger(1), NewInteger(2), NewInteger(-2), NewInteger(1))

	// FormatCurrency("1234.50") matches FormatCurrency(1234.5) across en-US, pt-BR, de-DE, fr-FR.
	localeTests := []struct {
		lcid      int
		localized string
		invariant string
	}{
		{1033, "1,234.50", "1234.50"},
		{1046, "1.234,50", "1234.50"},
		{1031, "1.234,50", "1234.50"},
		{1036, "1 234,50", "1234.50"},
	}
	for _, lt := range localeTests {
		host.Session().SetLCID(lt.lcid)
		expected := callBuiltin(t, vm, "FormatCurrency", NewDouble(1234.5))
		gotLoc := callBuiltin(t, vm, "FormatCurrency", NewString(lt.localized))
		gotInv := callBuiltin(t, vm, "FormatCurrency", NewString(lt.invariant))
		if gotLoc.Type != VTString || gotLoc.Str != expected.Str {
			t.Fatalf("LCID %d FormatCurrency(%q): expected %q, got %q", lt.lcid, lt.localized, expected.Str, gotLoc.Str)
		}
		if gotInv.Type != VTString || gotInv.Str != expected.Str {
			t.Fatalf("LCID %d FormatCurrency(%q): expected %q, got %q", lt.lcid, lt.invariant, expected.Str, gotInv.Str)
		}
	}
}

func TestBuiltinFormatFullVMExecutionAndErrors(t *testing.T) {
	script := `<%
On Error Resume Next
Session.LCID = 1033
Dim a, b, c, d, e1, e2
a = FormatCurrency("1234.50")
b = IsNull(FormatCurrency(Null))
c = FormatNumber("1.235e1", 1, , vbTrue, vbFalse)
d = FormatPercent("-0.00001", 2)

Err.Clear
Dim bad1
bad1 = FormatCurrency("not-a-number")
e1 = Err.Number

Err.Clear
Dim bad2
bad2 = FormatNumber(10, -2)
e2 = Err.Number

Response.Write a & "|" & CStr(b) & "|" & c & "|" & d & "|" & CStr(e1) & "|" & CStr(e2)
%>`
	compiler := NewASPCompiler(script)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	vm.SetHost(host)
	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	got := output.String()
	expected := "$1,234.50|True|12.4|0.00%|" +
		strconv.Itoa(vbscript.HRESULTFromVBScriptCode(vbscript.TypeMismatch)) + "|" +
		strconv.Itoa(vbscript.HRESULTFromVBScriptCode(vbscript.InvalidProcedureCallOrArgument))
	if got != expected {
		t.Fatalf("unexpected VM output: got %q, want %q", got, expected)
	}
}

// TestBuiltinCDateParsesPortugueseFormat verifies CDate handles Portuguese date strings.
func TestBuiltinCDateParsesPortugueseFormat(t *testing.T) {
	vm := NewVM(nil, nil, 5)
	host := NewMockHost()
	vm.SetHost(host)
	host.Session().SetLCID(int(PortugueseBrazil))

	// PT-BR uses day/month/year format
	result := callBuiltin(t, vm, "CDate", NewString("09/04/2026"))
	if result.Type != VTDate {
		t.Fatalf("expected VTDate from CDate, got %#v", result)
	}

	// Verify it parsed as day=9, month=4 (April), year=2026
	dateValue := time.Unix(0, result.Num).UTC()
	if dateValue.Day() != 9 || dateValue.Month() != time.April || dateValue.Year() != 2026 {
		t.Fatalf("expected April 9, 2026, got %v", dateValue)
	}
}

// TestBuiltinCDateParsesUSFormat verifies CDate handles US date strings.
func TestBuiltinCDateParsesUSFormat(t *testing.T) {
	vm := NewVM(nil, nil, 5)
	host := NewMockHost()
	vm.SetHost(host)
	host.Session().SetLCID(int(EnglishUS))

	// US uses month/day/year format
	result := callBuiltin(t, vm, "CDate", NewString("04/09/2026"))
	if result.Type != VTDate {
		t.Fatalf("expected VTDate from CDate, got %#v", result)
	}

	// Verify it parsed as month=4 (April), day=9, year=2026
	dateValue := time.Unix(0, result.Num).UTC()
	if dateValue.Day() != 9 || dateValue.Month() != time.April || dateValue.Year() != 2026 {
		t.Fatalf("expected April 9, 2026, got %v", dateValue)
	}
}

// TestBuiltinDateValueParses verifies DateValue parses locale-aware date strings.
func TestBuiltinDateValueParses(t *testing.T) {
	vm := NewVM(nil, nil, 5)
	host := NewMockHost()
	vm.SetHost(host)
	host.Session().SetLCID(int(PortugueseBrazil))

	// Portuguese long date
	result := callBuiltin(t, vm, "DateValue", NewString("09/04/2026"))
	if result.Type != VTDate {
		t.Fatalf("expected VTDate from DateValue, got %#v", result)
	}

	// Verify parsing respects PT-BR day/month/year order
	dateValue := time.Unix(0, result.Num).UTC()
	if dateValue.Day() != 9 || dateValue.Month() != time.April {
		t.Fatalf("expected day=9, month=April, got day=%d, month=%v", dateValue.Day(), dateValue.Month())
	}
}

// TestBuiltinTimeValueParses verifies TimeValue parses time strings in 24-hour format.
func TestBuiltinTimeValueParses(t *testing.T) {
	vm := NewVM(nil, nil, 5)
	host := NewMockHost()
	vm.SetHost(host)

	// 24-hour format time - parse in server timezone
	result := callBuiltin(t, vm, "TimeValue", NewString("13:45:30"))
	if result.Type != VTDate {
		t.Fatalf("expected VTDate from TimeValue, got %#v", result)
	}

	timeValue := time.Unix(0, result.Num).In(builtinCurrentLocation(vm))
	if timeValue.Hour() != 13 || timeValue.Minute() != 45 || timeValue.Second() != 30 {
		t.Fatalf("expected 13:45:30, got %02d:%02d:%02d", timeValue.Hour(), timeValue.Minute(), timeValue.Second())
	}
}

// TestBuiltinSessionLCIDAffectsFormatting verifies changing Session LCID affects formatting.
func TestBuiltinSessionLCIDAffectsFormatting(t *testing.T) {
	vm := NewVM(nil, nil, 5)
	host := NewMockHost()
	vm.SetHost(host)

	// First format with Portuguese Brazil
	host.Session().SetLCID(int(PortugueseBrazil))
	resultPT := callBuiltin(t, vm, "FormatNumber", NewDouble(1234.5), NewInteger(2))
	if resultPT.Str != "1.234,50" {
		t.Fatalf("expected PT-BR FormatNumber='1.234,50', got %q", resultPT.Str)
	}

	// Change to German and reformat
	host.Session().SetLCID(int(GermanGermany))
	resultDE := callBuiltin(t, vm, "FormatNumber", NewDouble(1234.5), NewInteger(2))
	if resultDE.Str != "1.234,50" {
		t.Fatalf("expected DE FormatNumber='1.234,50', got %q", resultDE.Str)
	}

	// Change to US and reformat
	host.Session().SetLCID(int(EnglishUS))
	resultUS := callBuiltin(t, vm, "FormatNumber", NewDouble(1234.5), NewInteger(2))
	if resultUS.Str != "1,234.50" {
		t.Fatalf("expected US FormatNumber='1,234.50', got %q", resultUS.Str)
	}
}

// TestBuiltinCurrencyFormattingVariants verifies currency symbols for specific locales.
func TestBuiltinCurrencyFormattingVariants(t *testing.T) {
	vm := NewVM(nil, nil, 5)
	host := NewMockHost()
	vm.SetHost(host)

	tests := []struct {
		lcid     MSLCID
		expected string
	}{
		{EnglishUS, "$1,234.50"},          // US Dollar
		{EnglishCanada, "$1,234.50"},      // Canadian Dollar
		{EnglishAustralia, "$1,234.50"},   // Australian Dollar
		{EnglishIndia, "₹1,234.50"},       // Indian Rupee
		{PortugueseBrazil, "R$ 1.234,50"}, // Brazilian Real
		{SpanishMexico, "$1,234.50"},      // Mexican Peso (USD commonly used)
		{SpanishArgentina, "$1.234,50"},   // Argentine Peso
		{GermanGermany, "€ 1.234,50"},     // Euro with German separator spacing
		{FrenchFrance, "€ 1 234,50"},      // Euro with French separators
	}

	for _, test := range tests {
		host.Session().SetLCID(int(test.lcid))
		result := callBuiltin(t, vm, "FormatCurrency", NewDouble(1234.5))
		if result.Type != VTString || result.Str != test.expected {
			t.Fatalf("LCID %d: expected %q, got %q", test.lcid, test.expected, result.Str)
		}
	}
}

// TestBuiltinMonthNameLocalization verifies MonthName returns locale-aware names.
func TestBuiltinMonthNameLocalization(t *testing.T) {
	vm := NewVM(nil, nil, 5)
	host := NewMockHost()
	vm.SetHost(host)

	// Test Portuguese month name
	host.Session().SetLCID(int(PortugueseBrazil))
	aprilPT := callBuiltin(t, vm, "MonthName", NewInteger(4), NewInteger(0))
	if aprilPT.Type != VTString || aprilPT.Str != "abril" {
		t.Fatalf("expected Portuguese MonthName(4)='abril', got %q", aprilPT.Str)
	}

	// Test abbreviated Portuguese month name
	aprilPTAbbr := callBuiltin(t, vm, "MonthName", NewInteger(4), NewInteger(1))
	if aprilPTAbbr.Type != VTString || aprilPTAbbr.Str != "abr" {
		t.Fatalf("expected Portuguese MonthName(4,1)='abr', got %q", aprilPTAbbr.Str)
	}

	// Switch to English
	host.Session().SetLCID(int(EnglishUS))
	aprilEN := callBuiltin(t, vm, "MonthName", NewInteger(4), NewInteger(0))
	if aprilEN.Type != VTString || aprilEN.Str != "April" {
		t.Fatalf("expected English MonthName(4)='April', got %q", aprilEN.Str)
	}
}

// TestBuiltinWeekdayNameLocalization verifies WeekdayName respects locale.
func TestBuiltinWeekdayNameLocalization(t *testing.T) {
	vm := NewVM(nil, nil, 5)
	host := NewMockHost()
	vm.SetHost(host)

	// Test Portuguese weekday name (Sunday = 1)
	host.Session().SetLCID(int(PortugueseBrazil))
	sundayPT := callBuiltin(t, vm, "WeekdayName", NewInteger(1), NewInteger(0))
	if sundayPT.Type != VTString || sundayPT.Str != "domingo" {
		t.Fatalf("expected Portuguese WeekdayName(1)='domingo', got %q", sundayPT.Str)
	}

	mondayPT := callBuiltin(t, vm, "WeekdayName", NewInteger(2), NewInteger(0))
	if mondayPT.Type != VTString || mondayPT.Str != "segunda-feira" {
		t.Fatalf("expected Portuguese WeekdayName(2)='segunda-feira', got %q", mondayPT.Str)
	}

	// English weekday name
	host.Session().SetLCID(int(EnglishUS))
	sundayEN := callBuiltin(t, vm, "WeekdayName", NewInteger(1), NewInteger(0))
	if sundayEN.Type != VTString || sundayEN.Str != "Sunday" {
		t.Fatalf("expected English WeekdayName(1)='Sunday', got %q", sundayEN.Str)
	}
}

// TestBuiltinDateStringParsing verifies composite date/time parsing across locales.
func TestBuiltinDateStringParsing(t *testing.T) {
	vm := NewVM(nil, nil, 5)
	host := NewMockHost()
	vm.SetHost(host)

	// Test with multiple date formats
	testFormats := []struct {
		lcid  MSLCID
		input string
		day   int
		month time.Month
		year  int
	}{
		{PortugueseBrazil, "09/04/2026", 9, time.April, 2026},

		{EnglishUS, "04/09/2026", 9, time.April, 2026},
		{GermanGermany, "09.04.2026", 9, time.April, 2026},
		{FrenchFrance, "09/04/2026", 9, time.April, 2026},
	}

	for _, test := range testFormats {
		host.Session().SetLCID(int(test.lcid))
		result := callBuiltin(t, vm, "CDate", NewString(test.input))
		if result.Type != VTDate {
			t.Fatalf("LCID %d: expected VTDate from CDate(%q), got %#v", test.lcid, test.input, result)
		}

		dateValue := time.Unix(0, result.Num).UTC()
		if dateValue.Day() != test.day || dateValue.Month() != test.month || dateValue.Year() != test.year {
			t.Fatalf("LCID %d: expected %d-%02d-%d, got %d-%02d-%d", test.lcid, test.year, test.month, test.day, dateValue.Year(), dateValue.Month(), dateValue.Day())
		}
	}
}

// TestBuiltinTimeFormatting verifies that Time() returns only the time part
// and Date() returns only the date part when converted to string.
func TestBuiltinTimeFormatting(t *testing.T) {
	vm := NewVM(nil, nil, 5)
	host := NewMockHost()
	vm.SetHost(host)
	host.Session().SetLCID(int(EnglishUS))

	// Time() should return only time (since it's anchored to 1899-12-30)
	timeVal := callBuiltin(t, vm, "Time")
	timeStr := vm.valueToString(timeVal)
	if strings.Contains(timeStr, "1899") {
		t.Fatalf("expected Time() string to omit date part, got %q", timeStr)
	}

	// Date() should return only date (since time is 00:00:00)
	dateVal := callBuiltin(t, vm, "Date")
	dateStr := vm.valueToString(dateVal)
	if strings.Contains(dateStr, ":") {
		t.Fatalf("expected Date() string to omit time part, got %q", dateStr)
	}

	// Now() should return both
	nowVal := callBuiltin(t, vm, "Now")
	nowStr := vm.valueToString(nowVal)
	if !strings.Contains(nowStr, "/") || !strings.Contains(nowStr, ":") {
		t.Fatalf("expected Now() string to include both date and time, got %q", nowStr)
	}
}

// TestBuiltinDateStringLocaleFormatting verifies that VTDate values are converted
// to strings using the active LCID locale settings via both vm.valueToString
// and the generic Value.String() method.
func TestBuiltinDateStringLocaleFormatting(t *testing.T) {
	// Create a fixed date: March 4, 2026 14:30:00 UTC
	testDate := time.Date(2026, time.March, 4, 14, 30, 0, 0, time.UTC)
	dateVal := NewDate(testDate)

	// Verify Value.String() uses locale-aware formatting (not RFC3339).
	// The default LCID in config is 1033 (EnglishUS).
	rawStr := dateVal.String()
	if rawStr == "" {
		t.Fatal("Value.String() on VTDate returned empty")
	}
	if strings.Contains(rawStr, "T") && strings.Contains(rawStr, "Z") {
		t.Fatalf("Value.String() should NOT return RFC3339 for VTDate, got %q", rawStr)
	}
	// US locale should format as "3/4/2026 2:30:00 PM" (or similar)
	if !strings.Contains(rawStr, "2026") {
		t.Fatalf("Value.String() for date should contain the year, got %q", rawStr)
	}

	// Verify CStr(date) uses locale-aware formatting
	vm := NewVM(nil, nil, 5)
	host := NewMockHost()
	vm.SetHost(host)
	host.Session().SetLCID(int(EnglishUS))

	cstr := callBuiltin(t, vm, "CStr", dateVal)
	if cstr.Type != VTString {
		t.Fatalf("expected VTString from CStr(date), got %#v", cstr)
	}
	if strings.Contains(cstr.Str, "T") && strings.Contains(cstr.Str, "Z") {
		t.Fatalf("CStr(date) should not return RFC3339, got %q", cstr.Str)
	}
	if !strings.Contains(cstr.Str, "2026") {
		t.Fatalf("CStr(date) should contain the year, got %q", cstr.Str)
	}
}

// TestBuiltinDateStringAcrossLocales verifies date-to-string conversion
// produces different formats for different LCID settings.
func TestBuiltinDateStringAcrossLocales(t *testing.T) {
	testDate := time.Date(2026, time.March, 4, 14, 30, 0, 0, time.UTC)
	dateVal := NewDate(testDate)

	tests := []struct {
		name        string
		lcid        MSLCID
		contains    string // substring that must appear in the formatted output
		notContains string // substring that must NOT appear
	}{
		{
			name:        "EnglishUS",
			lcid:        EnglishUS,
			contains:    "3/4/2026",
			notContains: "RFC3339",
		},
		{
			name:        "PortugueseBrazil",
			lcid:        PortugueseBrazil,
			contains:    "2026",
			notContains: "T",
		},
		{
			name:        "GermanGermany",
			lcid:        GermanGermany,
			contains:    "2026",
			notContains: "T",
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			vm := NewVM(nil, nil, 5)
			host := NewMockHost()
			vm.SetHost(host)
			host.Session().SetLCID(int(tt.lcid))

			result := vm.valueToString(dateVal)
			if result == "" {
				t.Fatal("vm.valueToString(date) returned empty")
			}
			if !strings.Contains(result, tt.contains) {
				t.Errorf("expected date string to contain %q, got %q", tt.contains, result)
			}
			if tt.notContains != "" && strings.Contains(result, tt.notContains) {
				t.Errorf("date string should NOT contain %q, got %q", tt.notContains, result)
			}
		})
	}
}

// TestBuiltinDateOnlyString verifies that date-only values (no time component)
// are formatted with only the date part, not the time part.
func TestBuiltinDateOnlyString(t *testing.T) {
	vm := NewVM(nil, nil, 5)
	host := NewMockHost()
	vm.SetHost(host)
	host.Session().SetLCID(int(EnglishUS))

	// DateSerial creates a date with no time component (midnight)
	dateOnly := callBuiltin(t, vm, "DateSerial", NewInteger(2026), NewInteger(3), NewInteger(4))
	dateStr := vm.valueToString(dateOnly)

	// Should contain date but not time separators
	if !strings.Contains(dateStr, "2026") {
		t.Fatalf("expected date-only string to contain year, got %q", dateStr)
	}
	// A pure date (midnight) should omit the time part
	if strings.Contains(dateStr, ":") {
		t.Fatalf("expected date-only string to omit time, got %q", dateStr)
	}

	// TimeSerial creates a time anchored to 1899-12-30
	timeOnly := callBuiltin(t, vm, "TimeSerial", NewInteger(14), NewInteger(30), NewInteger(0))
	timeStr := vm.valueToString(timeOnly)

	// Should NOT contain 1899 (base date should be hidden)
	if strings.Contains(timeStr, "1899") {
		t.Fatalf("expected time-only string to omit base date, got %q", timeStr)
	}
	if !strings.Contains(timeStr, ":") {
		t.Fatalf("expected time-only string to contain time, got %q", timeStr)
	}
}

// TestBuiltinImplicitDateStringCoercion verifies that implicit string coercion
// of dates (via concatenation) uses locale-aware formatting.
func TestBuiltinImplicitDateStringCoercion(t *testing.T) {
	testDate := time.Date(2026, time.March, 4, 14, 30, 0, 0, time.UTC)
	dateVal := NewDate(testDate)

	vm := NewVM(nil, nil, 5)
	host := NewMockHost()
	vm.SetHost(host)
	host.Session().SetLCID(int(EnglishUS))

	// Simulate string concatenation: date & ""
	concatResult := vm.concatValues(dateVal, NewString(""))
	if concatResult.Type != VTString {
		t.Fatalf("expected VTString from date concatenation, got %#v", concatResult)
	}
	if strings.Contains(concatResult.Str, "T") && strings.Contains(concatResult.Str, "Z") {
		t.Fatalf("date concatenation should not use RFC3339, got %q", concatResult.Str)
	}
	if !strings.Contains(concatResult.Str, "2026") {
		t.Fatalf("date concatenation should contain year, got %q", concatResult.Str)
	}
}
