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
	"strings"
	"testing"
	"time"

	"g3pix.com.br/axonasp/vbscript"
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

	configured := callBuiltin(t, vm, "GetLocale")
	if configured.Type != VTInteger || configured.Num != 1046 {
		t.Fatalf("expected configured default LCID 1046, got %#v", configured)
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
	if longTime.Type != VTString || longTime.Str != "13:05:07" {
		t.Fatalf("expected 24-hour long time in server timezone, got %#v", longTime)
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

// TestBuiltinReplaceNullExpression verifies Null propagation for Replace(expression,...).
func TestBuiltinReplaceNullExpression(t *testing.T) {
	vm := NewVM(nil, nil, 5)

	result := callBuiltin(t, vm, "Replace", NewNull(), NewString("a"), NewString("b"))
	if result.Type != VTNull {
		t.Fatalf("expected VTNull for Replace(Null,...), got %#v", result)
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
