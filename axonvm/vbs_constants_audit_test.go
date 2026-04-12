/*
 * AxonASP Server
 * Copyright (C) 2026 G3pix Ltda. All rights reserved.
 *
 * Developed by Lucas GuimarÃ£es - G3pix Ltda
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
)

// constantExpectation describes one expected VBScript constant entry from the Microsoft VBScript catalog.
type constantExpectation struct {
	Name string
	Kind ValueType
	Int  int64
	Text string
}

// expectedVBSConstantCatalog returns the full ordered constant catalog expected in VBSConstants.
func expectedVBSConstantCatalog() []constantExpectation {
	return []constantExpectation{
		{Name: "vbCr", Kind: VTString, Text: "\r"},
		{Name: "vbLf", Kind: VTString, Text: "\n"},
		{Name: "vbCrLf", Kind: VTString, Text: "\r\n"},
		{Name: "vbNewLine", Kind: VTString, Text: "\r\n"},
		{Name: "vbNullChar", Kind: VTString, Text: "\x00"},
		{Name: "vbNullString", Kind: VTString, Text: ""},
		{Name: "vbTab", Kind: VTString, Text: "\t"},
		{Name: "vbBack", Kind: VTString, Text: "\x08"},
		{Name: "vbFormFeed", Kind: VTString, Text: "\x0C"},
		{Name: "vbVerticalTab", Kind: VTString, Text: "\x0B"},
		{Name: "vbTrue", Kind: VTInteger, Int: -1},
		{Name: "vbFalse", Kind: VTInteger, Int: 0},
		{Name: "vbEmpty", Kind: VTInteger, Int: 0},
		{Name: "vbNull", Kind: VTInteger, Int: 1},
		{Name: "vbInteger", Kind: VTInteger, Int: 2},
		{Name: "vbLong", Kind: VTInteger, Int: 3},
		{Name: "vbSingle", Kind: VTInteger, Int: 4},
		{Name: "vbDouble", Kind: VTInteger, Int: 5},
		{Name: "vbCurrency", Kind: VTInteger, Int: 6},
		{Name: "vbDate", Kind: VTInteger, Int: 7},
		{Name: "vbString", Kind: VTInteger, Int: 8},
		{Name: "vbObject", Kind: VTInteger, Int: 9},
		{Name: "vbError", Kind: VTInteger, Int: 10},
		{Name: "vbBoolean", Kind: VTInteger, Int: 11},
		{Name: "vbVariant", Kind: VTInteger, Int: 12},
		{Name: "vbDataObject", Kind: VTInteger, Int: 13},
		{Name: "vbDecimal", Kind: VTInteger, Int: 14},
		{Name: "vbByte", Kind: VTInteger, Int: 17},
		{Name: "vbArray", Kind: VTInteger, Int: 8192},
		{Name: "vbGeneralDate", Kind: VTInteger, Int: 0},
		{Name: "vbLongDate", Kind: VTInteger, Int: 1},
		{Name: "vbShortDate", Kind: VTInteger, Int: 2},
		{Name: "vbLongTime", Kind: VTInteger, Int: 3},
		{Name: "vbShortTime", Kind: VTInteger, Int: 4},
		{Name: "vbBinaryCompare", Kind: VTInteger, Int: 0},
		{Name: "vbTextCompare", Kind: VTInteger, Int: 1},
		{Name: "vbDatabaseCompare", Kind: VTInteger, Int: 2},
		{Name: "vbUpperCase", Kind: VTInteger, Int: 1},
		{Name: "vbLowerCase", Kind: VTInteger, Int: 2},
		{Name: "vbProperCase", Kind: VTInteger, Int: 3},
		{Name: "vbWide", Kind: VTInteger, Int: 4},
		{Name: "vbNarrow", Kind: VTInteger, Int: 8},
		{Name: "vbKatakana", Kind: VTInteger, Int: 16},
		{Name: "vbHiragana", Kind: VTInteger, Int: 32},
		{Name: "vbUnicode", Kind: VTInteger, Int: 64},
		{Name: "vbFromUnicode", Kind: VTInteger, Int: 128},
		{Name: "vbSunday", Kind: VTInteger, Int: 1},
		{Name: "vbMonday", Kind: VTInteger, Int: 2},
		{Name: "vbTuesday", Kind: VTInteger, Int: 3},
		{Name: "vbWednesday", Kind: VTInteger, Int: 4},
		{Name: "vbThursday", Kind: VTInteger, Int: 5},
		{Name: "vbFriday", Kind: VTInteger, Int: 6},
		{Name: "vbSaturday", Kind: VTInteger, Int: 7},
		{Name: "vbJanuary", Kind: VTInteger, Int: 1},
		{Name: "vbFebruary", Kind: VTInteger, Int: 2},
		{Name: "vbMarch", Kind: VTInteger, Int: 3},
		{Name: "vbApril", Kind: VTInteger, Int: 4},
		{Name: "vbMay", Kind: VTInteger, Int: 5},
		{Name: "vbJune", Kind: VTInteger, Int: 6},
		{Name: "vbJuly", Kind: VTInteger, Int: 7},
		{Name: "vbAugust", Kind: VTInteger, Int: 8},
		{Name: "vbSeptember", Kind: VTInteger, Int: 9},
		{Name: "vbOctober", Kind: VTInteger, Int: 10},
		{Name: "vbNovember", Kind: VTInteger, Int: 11},
		{Name: "vbDecember", Kind: VTInteger, Int: 12},
		{Name: "vbUseSystemDayOfWeek", Kind: VTInteger, Int: 0},
		{Name: "vbUseSystem", Kind: VTInteger, Int: 0},
		{Name: "vbFirstJan1", Kind: VTInteger, Int: 1},
		{Name: "vbFirstFourDays", Kind: VTInteger, Int: 2},
		{Name: "vbFirstFullWeek", Kind: VTInteger, Int: 3},
		{Name: "vbBlack", Kind: VTInteger, Int: 0},
		{Name: "vbRed", Kind: VTInteger, Int: 16711680},
		{Name: "vbGreen", Kind: VTInteger, Int: 65280},
		{Name: "vbYellow", Kind: VTInteger, Int: 16776960},
		{Name: "vbBlue", Kind: VTInteger, Int: 255},
		{Name: "vbMagenta", Kind: VTInteger, Int: 16711935},
		{Name: "vbCyan", Kind: VTInteger, Int: 65535},
		{Name: "vbWhite", Kind: VTInteger, Int: 16777215},
		{Name: "vbOK", Kind: VTInteger, Int: 1},
		{Name: "vbCancel", Kind: VTInteger, Int: 2},
		{Name: "vbAbort", Kind: VTInteger, Int: 3},
		{Name: "vbRetry", Kind: VTInteger, Int: 4},
		{Name: "vbIgnore", Kind: VTInteger, Int: 5},
		{Name: "vbYes", Kind: VTInteger, Int: 6},
		{Name: "vbNo", Kind: VTInteger, Int: 7},
		{Name: "vbOKOnly", Kind: VTInteger, Int: 0},
		{Name: "vbOKCancel", Kind: VTInteger, Int: 1},
		{Name: "vbAbortRetryIgnore", Kind: VTInteger, Int: 2},
		{Name: "vbYesNoCancel", Kind: VTInteger, Int: 3},
		{Name: "vbYesNo", Kind: VTInteger, Int: 4},
		{Name: "vbRetryCancel", Kind: VTInteger, Int: 5},
		{Name: "vbCritical", Kind: VTInteger, Int: 16},
		{Name: "vbQuestion", Kind: VTInteger, Int: 32},
		{Name: "vbExclamation", Kind: VTInteger, Int: 48},
		{Name: "vbInformation", Kind: VTInteger, Int: 64},
		{Name: "vbDefaultButton1", Kind: VTInteger, Int: 0},
		{Name: "vbDefaultButton2", Kind: VTInteger, Int: 256},
		{Name: "vbDefaultButton3", Kind: VTInteger, Int: 512},
		{Name: "vbDefaultButton4", Kind: VTInteger, Int: 768},
		{Name: "vbApplicationModal", Kind: VTInteger, Int: 0},
		{Name: "vbSystemModal", Kind: VTInteger, Int: 4096},
		{Name: "vbMsgBoxHelpButton", Kind: VTInteger, Int: 16384},
		{Name: "vbMsgBoxSetForeground", Kind: VTInteger, Int: 65536},
		{Name: "vbMsgBoxRight", Kind: VTInteger, Int: 524288},
		{Name: "vbMsgBoxRtlReading", Kind: VTInteger, Int: 1048576},
		{Name: "vbObjectError", Kind: VTInteger, Int: -2147221504},
		{Name: "vbASCII", Kind: VTInteger, Int: 0},
		{Name: "vbUseDefault", Kind: VTInteger, Int: -2},
	}
}

// assertValueMatchesExpectation compares one runtime value with one expected constant entry.
func assertValueMatchesExpectation(t *testing.T, got Value, expected constantExpectation, context string) {
	t.Helper()

	if got.Type != expected.Kind {
		t.Fatalf("%s: type mismatch for %s: got=%v want=%v", context, expected.Name, got.Type, expected.Kind)
	}

	switch expected.Kind {
	case VTString:
		if got.Str != expected.Text {
			t.Fatalf("%s: string mismatch for %s: got=%q want=%q", context, expected.Name, got.Str, expected.Text)
		}
	default:
		if got.Num != expected.Int {
			t.Fatalf("%s: numeric mismatch for %s: got=%d want=%d", context, expected.Name, got.Num, expected.Int)
		}
	}
}

// TestVBScriptConstantCatalogIsExhaustive verifies the full VBScript catalog names, ordering,
// uniqueness, and values. The VBSConstants slice may be longer than the VBScript catalog because
// additional type library constants (e.g. ADODB ad-prefixed constants) are appended via init()
// in their respective files. This test only validates the VBScript-specific prefix entries.
func TestVBScriptConstantCatalogIsExhaustive(t *testing.T) {
	expected := expectedVBSConstantCatalog()
	if len(VBSConstants) < len(expected) {
		t.Fatalf("catalog too short: got=%d want>=%d", len(VBSConstants), len(expected))
	}

	seen := make(map[string]bool, len(VBSConstants))
	for index := range expected {
		entry := VBSConstants[index]
		exp := expected[index]

		if entry.Name != exp.Name {
			t.Fatalf("name mismatch at index %d: got=%q want=%q", index, entry.Name, exp.Name)
		}

		lower := strings.ToLower(entry.Name)
		if seen[lower] {
			t.Fatalf("duplicate constant name detected: %s", entry.Name)
		}
		seen[lower] = true

		assertValueMatchesExpectation(t, entry.Val, exp, "catalog")
	}
}

// TestVBScriptConstantCatalogInjectedIntoVM verifies VM global slots receive the same constant catalog values.
func TestVBScriptConstantCatalogInjectedIntoVM(t *testing.T) {
	compiler := NewASPCompiler(`<% %>`)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	expected := expectedVBSConstantCatalog()
	// Global layout: 7 intrinsic objects (0-6, including Err) + 2 transaction event handler slots (7-8) + builtins + constants
	constStart := 9 + len(BuiltinRegistry)

	for index := range expected {
		slot := constStart + index
		if slot >= len(vm.Globals) {
			t.Fatalf("constant global slot out of range: slot=%d globals=%d", slot, len(vm.Globals))
		}
		assertValueMatchesExpectation(t, vm.Globals[slot], expected[index], "vm globals")
	}
}
