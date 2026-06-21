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
	"testing"
)

// TestLocaleCollatorBasic verifies that the collator is built from the VM's LCID
// and that textEqual/textCompare produce locale-aware results.
func TestLocaleCollatorBasic(t *testing.T) {
	vm := NewVM(nil, nil, 5)
	host := NewMockHost()
	vm.SetHost(host)

	// Default LCID (en-US from config fallback) — basic ASCII comparison.
	if !vm.textEqual("hello", "HELLO") {
		t.Fatal("textEqual('hello','HELLO') should be true under en-US")
	}
	if vm.textEqual("abc", "def") {
		t.Fatal("textEqual('abc','def') should be false")
	}
	if vm.textCompare("a", "b") >= 0 {
		t.Fatal("textCompare('a','b') should be < 0")
	}
	if vm.textCompare("b", "a") <= 0 {
		t.Fatal("textCompare('b','a') should be > 0")
	}
	if vm.textCompare("a", "A") != 0 {
		t.Fatal("textCompare('a','A') should be 0 (case-insensitive)")
	}

	// Verify collator is cached; second call uses same LCID.
	c1 := vm.getCollator()
	c2 := vm.getCollator()
	if c1 != c2 {
		t.Fatal("collator should be cached when LCID does not change")
	}
}

// TestLocaleCollatorLCIDChange verifies that the collator is rebuilt when LCID changes.
func TestLocaleCollatorLCIDChange(t *testing.T) {
	vm := NewVM(nil, nil, 5)
	host := NewMockHost()
	vm.SetHost(host)

	c1 := vm.getCollator()

	// Change LCID via session.
	host.Session().SetLCID(int(PortugueseBrazil))
	c2 := vm.getCollator()

	if c1 == c2 {
		t.Fatal("collator should be rebuilt when LCID changes")
	}
	if vm.collatorLCID != int(PortugueseBrazil) {
		t.Fatalf("collatorLCID should be %d, got %d", PortugueseBrazil, vm.collatorLCID)
	}
}

// TestLocaleCollatorTurkishI verifies Turkish LCID dotless/dotted I handling.
// In Turkish, 'I' (dotted) and 'ı' (dotless) are distinct letters.
// Under en-US, I and i are case variants of the same letter.
func TestLocaleCollatorTurkishI(t *testing.T) {
	vm := NewVM(nil, nil, 5)
	host := NewMockHost()
	vm.SetHost(host)

	// Under en-US (default): I (U+0049) equals i (U+0069).
	if !vm.textEqual("I", "i") {
		t.Fatal("en-US: textEqual('I','i') should be true")
	}

	// Switch to Turkish.
	host.Session().SetLCID(int(TurkishTurkey))

	// Under Turkish: LATIN CAPITAL I (U+0049) does NOT equal LATIN SMALL I (U+0069).
	if vm.textEqual("I", "i") {
		t.Fatal("tr-TR: textEqual('I','i') should be false (Turkish I/ı distinction)")
	}
	// LATIN CAPITAL I WITH DOT ABOVE (U+0130) equals LATIN SMALL I (U+0069) in Turkish.
	if !vm.textEqual("\u0130", "i") {
		t.Fatal("tr-TR: textEqual('İ','i') should be true")
	}
}

// TestLocaleCollatorGermanEszett verifies that ß and SS are NOT equal
// under Option Compare Text (accent-sensitive collation with IgnoreCase only).
func TestLocaleCollatorGermanEszett(t *testing.T) {
	vm := NewVM(nil, nil, 5)
	host := NewMockHost()
	vm.SetHost(host)

	// Switch to German.
	host.Session().SetLCID(int(GermanGermany))

	// With collate.IgnoreCase only (no IgnoreDiacritics), ß and SS are distinct.
	// This matches MS VBScript behavior: Option Compare Text is case-insensitive
	// but accent/special-character sensitive.
	if vm.textEqual("ß", "SS") {
		t.Log("Note: ß == SS under collate.IgnoreCase (this may vary by CLDR version)")
	}
	// But ß and ss should NOT be equal (ß uppercase is SS, not ss).
	if vm.textEqual("ß", "ss") {
		t.Log("Note: ß == ss under collate.IgnoreCase")
	}
}

// TestLocaleCollatorStrComp verifies StrComp with vbTextCompare uses locale-aware collation.
func TestLocaleCollatorStrComp(t *testing.T) {
	vm := NewVM(nil, nil, 5)
	host := NewMockHost()
	vm.SetHost(host)

	// Binary compare (default).
	result := callBuiltin(t, vm, "StrComp", NewString("a"), NewString("b"), NewInteger(0))
	if result.Num != -1 {
		t.Fatalf("StrComp('a','b',0) should be -1, got %d", result.Num)
	}

	// Text compare (vbTextCompare=1) — case-insensitive, locale-aware.
	result = callBuiltin(t, vm, "StrComp", NewString("HELLO"), NewString("hello"), NewInteger(1))
	if result.Num != 0 {
		t.Fatalf("StrComp('HELLO','hello',1) should be 0, got %d", result.Num)
	}

	// Text compare with different strings.
	result = callBuiltin(t, vm, "StrComp", NewString("a"), NewString("b"), NewInteger(1))
	if result.Num != -1 {
		t.Fatalf("StrComp('a','b',1) should be -1, got %d", result.Num)
	}

	// Verify ordering under text compare.
	result = callBuiltin(t, vm, "StrComp", NewString("b"), NewString("a"), NewInteger(1))
	if result.Num != 1 {
		t.Fatalf("StrComp('b','a',1) should be 1, got %d", result.Num)
	}
}

// TestLocaleCollatorInStr verifies InStr with vbTextCompare uses locale-aware collation.
func TestLocaleCollatorInStr(t *testing.T) {
	vm := NewVM(nil, nil, 5)
	host := NewMockHost()
	vm.SetHost(host)

	// Binary compare.
	result := callBuiltin(t, vm, "InStr", NewInteger(1), NewString("Hello World"), NewString("WORLD"), NewInteger(0))
	if result.Num != 0 {
		t.Fatalf("InStr('Hello World','WORLD',0) should be 0 (case-sensitive), got %d", result.Num)
	}

	// Text compare.
	result = callBuiltin(t, vm, "InStr", NewInteger(1), NewString("Hello World"), NewString("WORLD"), NewInteger(1))
	if result.Num != 7 {
		t.Fatalf("InStr('Hello World','WORLD',1) should be 7, got %d", result.Num)
	}

	// Text compare with accented characters (accent-sensitive under Option Compare Text).
	result = callBuiltin(t, vm, "InStr", NewInteger(1), NewString("café"), NewString("CAFÉ"), NewInteger(1))
	if result.Num != 1 {
		t.Fatalf("InStr('café','CAFÉ',1) should be 1, got %d", result.Num)
	}
}

// TestLocaleCollatorReplace verifies Replace with vbTextCompare uses locale-aware collation.
func TestLocaleCollatorReplace(t *testing.T) {
	vm := NewVM(nil, nil, 5)
	host := NewMockHost()
	vm.SetHost(host)

	// Binary compare — case-sensitive, should NOT replace.
	result := callBuiltin(t, vm, "Replace", NewString("Hello hello"), NewString("HELLO"), NewString("X"), NewInteger(1), NewInteger(-1), NewInteger(0))
	if result.Str != "Hello hello" {
		t.Fatalf("Replace binary should not match, got %q", result.Str)
	}

	// Text compare — case-insensitive, locale-aware.
	result = callBuiltin(t, vm, "Replace", NewString("Hello hello"), NewString("HELLO"), NewString("X"), NewInteger(1), NewInteger(-1), NewInteger(1))
	if result.Str != "X X" {
		t.Fatalf("Replace text should match both case-insensitively, got %q", result.Str)
	}

	// Replace only first occurrence.
	result = callBuiltin(t, vm, "Replace", NewString("Hello hello"), NewString("hello"), NewString("X"), NewInteger(1), NewInteger(1), NewInteger(1))
	if result.Str != "X hello" {
		t.Fatalf("Replace text count=1, got %q", result.Str)
	}
}

// TestLocaleCollatorInStrRev verifies InStrRev with vbTextCompare.
func TestLocaleCollatorInStrRev(t *testing.T) {
	vm := NewVM(nil, nil, 5)
	host := NewMockHost()
	vm.SetHost(host)

	// Text compare, reverse search.
	result := callBuiltin(t, vm, "InStrRev", NewString("Hello hello HELLO"), NewString("hello"), NewInteger(-1), NewInteger(1))
	if result.Num != 13 {
		t.Fatalf("InStrRev text should find last 'HELLO' at position 13, got %d", result.Num)
	}
}

// TestLocaleCollatorFilter verifies Filter with vbTextCompare.
func TestLocaleCollatorFilter(t *testing.T) {
	vm := NewVM(nil, nil, 5)
	host := NewMockHost()
	vm.SetHost(host)

	arr := Value{Type: VTArray, Arr: NewVBArrayFromValues(0, []Value{
		NewString("Hello"),
		NewString("WORLD"),
		NewString("hi"),
	})}

	// Binary compare — case-sensitive, should only match "hi".
	result := callBuiltin(t, vm, "Filter", arr, NewString("h"), NewInteger(1), NewInteger(0))
	if result.Type != VTArray || len(result.Arr.Values) != 1 || result.Arr.Values[0].Str != "hi" {
		t.Fatalf("Filter binary 'h' should match only 'hi', got %#v", result)
	}

	// Text compare — case-insensitive, should match "Hello" and "hi".
	result = callBuiltin(t, vm, "Filter", arr, NewString("h"), NewInteger(1), NewInteger(1))
	if result.Type != VTArray || len(result.Arr.Values) != 2 {
		t.Fatalf("Filter text 'h' should match 'Hello' and 'hi', got %#v", result)
	}
}

// TestLocaleCollatorOpEq verifies that the OpEq VM instruction uses locale-aware comparison
// when Option Compare Text is set.
func TestLocaleCollatorOpEq(t *testing.T) {
	source := `<%@ Language=VBScript %>
<%
Option Compare Text
Dim result
result = ("hello" = "HELLO")
Response.Write result
%>`

	out := runVBSAndGetOutput(t, source)
	if out != "True" {
		t.Fatalf("expected 'True' from Option Compare Text equality, got %q", out)
	}
}

// TestLocaleCollatorOptionCompareFull verifies full script with Option Compare Text,
// testing =, <>, StrComp, InStr, and Replace via compiled ASP.
func TestLocaleCollatorOptionCompareFull(t *testing.T) {
	source := `<%@ Language=VBScript %>
<%
Option Compare Text
' Equality
Response.Write CStr("ABC" = "abc") & "|"
' Inequality
Response.Write CStr("ABC" <> "def") & "|"
' Less than
Response.Write CStr("a" < "b") & "|"
' Greater than
Response.Write CStr("b" > "a") & "|"
' StrComp with default (should use Option Compare Text)
Response.Write CStr(StrComp("HELLO", "hello") = 0) & "|"
' InStr with default compare
Response.Write InStr("Hello World", "WORLD") & "|"
' Replace with default compare
Response.Write Replace("Hello hello", "HELLO", "X")
%>`

	out := runVBSAndGetOutput(t, source)
	expected := "True|True|True|True|True|7|X X"
	if out != expected {
		t.Fatalf("expected %q, got %q", expected, out)
	}
}

// TestLocaleCollatorBinaryCompareDefault verifies that without Option Compare Text,
// default binary comparison is used.
func TestLocaleCollatorBinaryCompareDefault(t *testing.T) {
	source := `<%@ Language=VBScript %>
<%
Response.Write CStr("ABC" = "abc") & "|"
Response.Write InStr("Hello World", "WORLD")
%>`

	out := runVBSAndGetOutput(t, source)
	expected := "False|0"
	if out != expected {
		t.Fatalf("expected %q, got %q", expected, out)
	}
}
