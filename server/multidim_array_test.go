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
package server

import (
	"strings"
	"testing"
)

// TestMultiDimensionalArrayAssignment tests that multi-dimensional array
// assignment (arr(0, 1) = value) works correctly by navigating through
// nested VBArrays instead of overwriting the inner VBArray.
func TestMultiDimensionalArrayAssignment(t *testing.T) {
	// Create a 2D array: ReDim arr(2, 3) -> 3 x 4 elements
	// In VBScript: arr(dim1, dim2)
	// Internally: VBArray[3] { VBArray[4], VBArray[4], VBArray[4] }
	outer := NewVBArray(0, 3)
	for i := 0; i < 3; i++ {
		outer.Values[i] = NewVBArray(0, 4)
	}

	// Test 1: Set arr(0, 2) = "hello"
	inner0, ok := toVBArray(outer.Values[0])
	if !ok {
		t.Fatal("inner array at index 0 should be a VBArray")
	}
	if !inner0.Set(2, "hello") {
		t.Fatal("Set(2, 'hello') should succeed")
	}

	// Verify arr(0, 2) == "hello"
	val, exists := inner0.Get(2)
	if !exists {
		t.Fatal("Get(2) should exist")
	}
	if val != "hello" {
		t.Fatalf("Expected 'hello', got '%v'", val)
	}

	// Verify arr(0, 0) is still nil (not overwritten)
	val0, exists0 := inner0.Get(0)
	if !exists0 {
		t.Fatal("Get(0) should exist (even if nil)")
	}
	if val0 != nil {
		t.Fatalf("Expected nil at arr(0,0), got '%v'", val0)
	}

	// Test 2: Set arr(1, 3) = "world"
	inner1, ok := toVBArray(outer.Values[1])
	if !ok {
		t.Fatal("inner array at index 1 should be a VBArray")
	}
	if !inner1.Set(3, "world") {
		t.Fatal("Set(3, 'world') should succeed")
	}

	// Verify arr(1, 3) == "world"
	val, exists = inner1.Get(3)
	if !exists {
		t.Fatal("Get(3) should exist")
	}
	if val != "world" {
		t.Fatalf("Expected 'world', got '%v'", val)
	}

	// Verify arr(0, 2) still == "hello" (not affected by arr(1, 3) assignment)
	inner0Again, ok := toVBArray(outer.Values[0])
	if !ok {
		t.Fatal("inner array at index 0 should still be a VBArray")
	}
	val, exists = inner0Again.Get(2)
	if !exists || val != "hello" {
		t.Fatalf("Expected 'hello' at arr(0,2), got '%v'", val)
	}

	// Test 3: Multi-dimensional read via standard navigation
	current := interface{}(outer)
	// Read arr(1, 3)
	currArr, ok := toVBArray(current)
	if !ok {
		t.Fatal("outer should be a VBArray")
	}
	innerVal, exists := currArr.Get(1)
	if !exists {
		t.Fatal("Get(1) should exist on outer")
	}
	innerArr, ok := toVBArray(innerVal)
	if !ok {
		t.Fatal("inner should be a VBArray")
	}
	finalVal, exists := innerArr.Get(3)
	if !exists {
		t.Fatal("Get(3) should exist on inner")
	}
	if finalVal != "world" {
		t.Fatalf("Expected 'world', got '%v'", finalVal)
	}
}

// TestRegExpMatchesEnumeration tests that RegExpMatchesCollection
// supports the Enumeration() interface for For Each iteration.
func TestRegExpMatchesEnumeration(t *testing.T) {
	re := &G3REGEXP{}
	re.SetProperty("pattern", `\d+`)
	re.SetProperty("global", true)

	result := re.Execute("abc 123 def 456 ghi")
	matches, ok := result.(*RegExpMatchesCollection)
	if !ok {
		t.Fatal("Execute should return *RegExpMatchesCollection")
	}

	if matches.count != 2 {
		t.Fatalf("Expected 2 matches, got %d", matches.count)
	}

	// Test Enumeration() interface
	enumerable, ok := result.(interface{ Enumeration() []interface{} })
	if !ok {
		t.Fatal("RegExpMatchesCollection should implement Enumeration() interface")
	}

	items := enumerable.Enumeration()
	if len(items) != 2 {
		t.Fatalf("Enumeration should return 2 items, got %d", len(items))
	}

	// Verify match values
	match0, ok := items[0].(*RegExpMatch)
	if !ok {
		t.Fatal("Item 0 should be *RegExpMatch")
	}
	if match0.Value != "123" {
		t.Fatalf("Expected '123', got '%s'", match0.Value)
	}

	match1, ok := items[1].(*RegExpMatch)
	if !ok {
		t.Fatal("Item 1 should be *RegExpMatch")
	}
	if match1.Value != "456" {
		t.Fatalf("Expected '456', got '%s'", match1.Value)
	}
}

// TestInStrCompareMode tests that InStr respects the compare parameter.
func TestInStrCompareMode(t *testing.T) {
	// Test case-sensitive (binary compare, default)
	// InStr(1, "Hello World", "hello") should return 0 (not found)
	result, handled := EvalBuiltInFunction("instr", []interface{}{1, "Hello World", "hello"}, nil)
	if !handled {
		t.Fatal("InStr should be handled")
	}
	if toInt(result) != 0 {
		t.Fatalf("Binary compare: InStr(1, 'Hello World', 'hello') should return 0, got %v", result)
	}

	// Test case-insensitive (text compare)
	// InStr(1, "Hello World", "hello", 1) should return 1
	result, handled = EvalBuiltInFunction("instr", []interface{}{1, "Hello World", "hello", 1}, nil)
	if !handled {
		t.Fatal("InStr should be handled")
	}
	if toInt(result) != 1 {
		t.Fatalf("Text compare: InStr(1, 'Hello World', 'hello', 1) should return 1, got %v", result)
	}

	// Test explicit binary compare
	// InStr(1, "Hello World", "hello", 0) should return 0
	result, handled = EvalBuiltInFunction("instr", []interface{}{1, "Hello World", "hello", 0}, nil)
	if !handled {
		t.Fatal("InStr should be handled")
	}
	if toInt(result) != 0 {
		t.Fatalf("Explicit binary compare: InStr(1, 'Hello World', 'hello', 0) should return 0, got %v", result)
	}

	// Test case-sensitive find (should work)
	// InStr(1, "Hello World", "Hello", 0) should return 1
	result, handled = EvalBuiltInFunction("instr", []interface{}{1, "Hello World", "Hello", 0}, nil)
	if !handled {
		t.Fatal("InStr should be handled")
	}
	if toInt(result) != 1 {
		t.Fatalf("Binary compare exact: InStr(1, 'Hello World', 'Hello', 0) should return 1, got %v", result)
	}

	// Test 2-arg form (default binary compare)
	// InStr("Hello World", "hello") should return 0
	result, handled = EvalBuiltInFunction("instr", []interface{}{"Hello World", "hello"}, nil)
	if !handled {
		t.Fatal("InStr should be handled")
	}
	if toInt(result) != 0 {
		t.Fatalf("2-arg default: InStr('Hello World', 'hello') should return 0, got %v", result)
	}

	// Test 2-arg form exact match
	// InStr("Hello World", "Hello") should return 1
	result, handled = EvalBuiltInFunction("instr", []interface{}{"Hello World", "Hello"}, nil)
	if !handled {
		t.Fatal("InStr should be handled")
	}
	if toInt(result) != 1 {
		t.Fatalf("2-arg exact: InStr('Hello World', 'Hello') should return 1, got %v", result)
	}
}

// TestReplaceCompareMode tests that Replace respects the compare parameter
// and that the start parameter correctly truncates output.
func TestReplaceCompareMode(t *testing.T) {
	// Test default (binary compare, case-sensitive)
	// Replace("Hello World", "hello", "HI") should return "Hello World" (no match)
	result, handled := EvalBuiltInFunction("replace", []interface{}{"Hello World", "hello", "HI"}, nil)
	if !handled {
		t.Fatal("Replace should be handled")
	}
	if toString(result) != "Hello World" {
		t.Fatalf("Binary default: Replace('Hello World', 'hello', 'HI') should return 'Hello World', got '%v'", result)
	}

	// Test case-insensitive (text compare)
	// Replace("Hello World", "hello", "HI", 1, -1, 1) should return "HI World"
	result, handled = EvalBuiltInFunction("replace", []interface{}{"Hello World", "hello", "HI", 1, -1, 1}, nil)
	if !handled {
		t.Fatal("Replace should be handled")
	}
	if toString(result) != "HI World" {
		t.Fatalf("Text compare: Replace('Hello World', 'hello', 'HI', 1, -1, 1) should return 'HI World', got '%v'", result)
	}

	// Test exact match (binary compare)
	// Replace("Hello World", "Hello", "HI") should return "HI World"
	result, handled = EvalBuiltInFunction("replace", []interface{}{"Hello World", "Hello", "HI"}, nil)
	if !handled {
		t.Fatal("Replace should be handled")
	}
	if toString(result) != "HI World" {
		t.Fatalf("Binary exact: Replace('Hello World', 'Hello', 'HI') should return 'HI World', got '%v'", result)
	}

	// Test start parameter: output should start from the start position
	// Replace("xxHello World", "Hello", "HI", 3) -> "HI World"
	result, handled = EvalBuiltInFunction("replace", []interface{}{"xxHello World", "Hello", "HI", 3}, nil)
	if !handled {
		t.Fatal("Replace should be handled")
	}
	if toString(result) != "HI World" {
		t.Fatalf("Start param: Replace('xxHello World', 'Hello', 'HI', 3) should return 'HI World', got '%v'", result)
	}

	// Test that [SITEINFO("name")] replacement pattern works
	// This mimics what QuickerSite does
	template := `Welcome to [SITEINFO("name")]!`
	result, handled = EvalBuiltInFunction("replace", []interface{}{template, `[SITEINFO("name")]`, "My Site"}, nil)
	if !handled {
		t.Fatal("Replace should be handled")
	}
	expected := "Welcome to My Site!"
	if toString(result) != expected {
		t.Fatalf("Template replace: expected '%s', got '%v'", expected, result)
	}
}

// TestInStrRevCompareMode tests that InStrRev respects the compare parameter.
func TestInStrRevCompareMode(t *testing.T) {
	// Default (binary compare)
	// InStrRev("Hello World Hello", "hello") should return 0 (case-sensitive, no match)
	result, handled := EvalBuiltInFunction("instrrev", []interface{}{"Hello World Hello", "hello"}, nil)
	if !handled {
		t.Fatal("InStrRev should be handled")
	}
	if toInt(result) != 0 {
		t.Fatalf("Binary default: InStrRev should return 0, got %v", result)
	}

	// Text compare
	// InStrRev("Hello World Hello", "hello", -1, 1) should return 13
	result, handled = EvalBuiltInFunction("instrrev", []interface{}{"Hello World Hello", "hello", -1, 1}, nil)
	if !handled {
		t.Fatal("InStrRev should be handled")
	}
	if toInt(result) != 13 {
		t.Fatalf("Text compare: InStrRev should return 13, got %v", result)
	}

	// Exact match (binary compare)
	// InStrRev("Hello World Hello", "Hello") should return 13
	result, handled = EvalBuiltInFunction("instrrev", []interface{}{"Hello World Hello", "Hello"}, nil)
	if !handled {
		t.Fatal("InStrRev should be handled")
	}
	if toInt(result) != 13 {
		t.Fatalf("Binary exact: InStrRev should return 13, got %v", result)
	}
}

// TestReplaceWithBracketedTags verifies that Replace correctly handles
// the bracket-tag patterns used by QuickerSite CMS (e.g., [MENU], [SITEINFO("name")]).
func TestReplaceWithBracketedTags(t *testing.T) {
	tests := []struct {
		name     string
		input    string
		find     string
		replace  string
		expected string
	}{
		{
			name:     "Simple bracket tag",
			input:    "Content [MENU] here",
			find:     "[MENU]",
			replace:  "<nav>Home | About</nav>",
			expected: "Content <nav>Home | About</nav> here",
		},
		{
			name:     "Bracket tag with quotes",
			input:    `Page: [SITEINFO("name")]`,
			find:     `[SITEINFO("name")]`,
			replace:  "My Website",
			expected: "Page: My Website",
		},
		{
			name:     "Multiple bracket tags",
			input:    `[SITENAME] - [SITESLOGAN]`,
			find:     "[SITENAME]",
			replace:  "AxonASP",
			expected: "AxonASP - [SITESLOGAN]",
		},
		{
			name:     "Case sensitivity matters",
			input:    `[sitename] and [SITENAME]`,
			find:     "[SITENAME]",
			replace:  "Test",
			expected: "[sitename] and Test",
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			result, handled := EvalBuiltInFunction("replace", []interface{}{tt.input, tt.find, tt.replace}, nil)
			if !handled {
				t.Fatal("Replace should be handled")
			}
			if toString(result) != tt.expected {
				t.Fatalf("Expected '%s', got '%v'", tt.expected, result)
			}
		})
	}
}

// TestRegExpReplaceWithBracketedTags tests RegExp Replace with bracket tag patterns
// similar to what QuickerSite's treatConstants function uses.
func TestRegExpReplaceWithBracketedTags(t *testing.T) {
	re := &G3REGEXP{}

	// Pattern similar to QuickerSite: match [CONSTNAME] or [CONSTNAME("param")]
	re.SetProperty("pattern", `\[SITEINFO\("([^"]+)"\)\]`)
	re.SetProperty("global", true)
	re.SetProperty("ignorecase", true)

	// Test matching
	result := re.Test(`Hello [SITEINFO("name")] World`)
	if !result {
		t.Fatal("Pattern should match [SITEINFO(\"name\")]")
	}

	// Test Execute returns matches
	matches := re.Execute(`Hello [SITEINFO("name")] and [SITEINFO("slogan")]`)
	matchCol, ok := matches.(*RegExpMatchesCollection)
	if !ok {
		t.Fatal("Execute should return *RegExpMatchesCollection")
	}
	if matchCol.count != 2 {
		t.Fatalf("Expected 2 matches, got %d", matchCol.count)
	}

	// Test Replace
	replaced := re.Replace(`[SITEINFO("name")] - [SITEINFO("slogan")]`, "REPLACED")
	if !strings.Contains(replaced, "REPLACED") {
		t.Fatalf("Replace should contain 'REPLACED', got '%s'", replaced)
	}
}
