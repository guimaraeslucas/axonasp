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
	"testing"
)

// runVBScriptExpr compiles a single VBScript expression and returns the resulting Value.
func runVBScriptExpr(t *testing.T, expr string) Value {
	t.Helper()
	code := "res = (" + expr + ")"
	compiler := NewCompiler(code)
	err := compiler.Compile()
	if err != nil {
		t.Fatalf("Compilation of %q failed: %v", expr, err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	vm.SetExecutionMode(ExecutionModeCLI)
	err = vm.Run()
	if err != nil {
		t.Fatalf("Execution of %q failed: %v", expr, err)
	}

	resIdx, ok := compiler.Globals.Get("res")
	if !ok {
		t.Fatalf("Global 'res' not found after executing %q", expr)
	}
	return vm.Globals[resIdx]
}

// TestTimeLiteralParsingRegression tests pure time literals, date literals,
// and combined date-time literals in VBScript to ensure compatibility with MS IIS/ASP VBScript.
func TestTimeLiteralParsingRegression(t *testing.T) {
	tests := []struct {
		name     string
		expr     string
		expected int64
	}{
		// Pure time literals - 24 hour format
		{name: "Hour(#13:45:30#)", expr: "Hour(#13:45:30#)", expected: 13},
		{name: "Minute(#13:45:30#)", expr: "Minute(#13:45:30#)", expected: 45},
		{name: "Second(#13:45:30#)", expr: "Second(#13:45:30#)", expected: 30},

		// Pure time literals - 12 hour PM format
		{name: "Hour(#1:45:30 PM#)", expr: "Hour(#1:45:30 PM#)", expected: 13},
		{name: "Minute(#1:45:30 PM#)", expr: "Minute(#1:45:30 PM#)", expected: 45},
		{name: "Second(#1:45:30 PM#)", expr: "Second(#1:45:30 PM#)", expected: 30},

		// Pure time literals - 12 hour AM format
		{name: "Hour(#1:45:30 AM#)", expr: "Hour(#1:45:30 AM#)", expected: 1},
		{name: "Minute(#1:45:30 AM#)", expr: "Minute(#1:45:30 AM#)", expected: 45},
		{name: "Second(#1:45:30 AM#)", expr: "Second(#1:45:30 AM#)", expected: 30},

		// Pure time literals without seconds
		{name: "Hour(#13:45#)", expr: "Hour(#13:45#)", expected: 13},
		{name: "Minute(#13:45#)", expr: "Minute(#13:45#)", expected: 45},
		{name: "Second(#13:45#)", expr: "Second(#13:45#)", expected: 0},

		// Date literals
		{name: "Month(#2024-06-15#)", expr: "Month(#2024-06-15#)", expected: 6},
		{name: "Day(#2024-06-15#)", expr: "Day(#2024-06-15#)", expected: 15},
		{name: "Year(#2024-06-15#)", expr: "Year(#2024-06-15#)", expected: 2024},

		// Combined date-time literals (ISO)
		{name: "Hour(#2024-06-15 13:45:30#)", expr: "Hour(#2024-06-15 13:45:30#)", expected: 13},
		{name: "Minute(#2024-06-15 13:45:30#)", expr: "Minute(#2024-06-15 13:45:30#)", expected: 45},
		{name: "Second(#2024-06-15 13:45:30#)", expr: "Second(#2024-06-15 13:45:30#)", expected: 30},
		{name: "Month(#2024-06-15 13:45:30#)", expr: "Month(#2024-06-15 13:45:30#)", expected: 6},
		{name: "Day(#2024-06-15 13:45:30#)", expr: "Day(#2024-06-15 13:45:30#)", expected: 15},
		{name: "Year(#2024-06-15 13:45:30#)", expr: "Year(#2024-06-15 13:45:30#)", expected: 2024},

		// Combined date-time literals (US format)
		{name: "Hour(#6/15/2024 1:45:30 PM#)", expr: "Hour(#6/15/2024 1:45:30 PM#)", expected: 13},
		{name: "Minute(#6/15/2024 1:45:30 PM#)", expr: "Minute(#6/15/2024 1:45:30 PM#)", expected: 45},
		{name: "Second(#6/15/2024 1:45:30 PM#)", expr: "Second(#6/15/2024 1:45:30 PM#)", expected: 30},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			val := runVBScriptExpr(t, tt.expr)
			if val.Type != VTInteger {
				t.Fatalf("expected VTInteger from %s, got type %v", tt.name, val.Type)
			}
			if val.Num != tt.expected {
				t.Errorf("%s = %d; expected %d", tt.name, val.Num, tt.expected)
			}
		})
	}
}

// TestColorConstantsBGRRegression tests all 8 built-in vb* color constants to verify exact BGR values
// according to Microsoft Script 5.6 documentation and prevent byte-order regressions.
func TestColorConstantsBGRRegression(t *testing.T) {
	colorTests := []struct {
		name     string
		hexStr   string
		expected int64
	}{
		{name: "vbBlack", hexStr: "&h00", expected: 0},
		{name: "vbRed", hexStr: "&hFF", expected: 255},
		{name: "vbGreen", hexStr: "&hFF00", expected: 65280},
		{name: "vbYellow", hexStr: "&hFFFF", expected: 65535},
		{name: "vbBlue", hexStr: "&hFF0000", expected: 16711680},
		{name: "vbMagenta", hexStr: "&hFF00FF", expected: 16711935},
		{name: "vbCyan", hexStr: "&hFFFF00", expected: 16776960},
		{name: "vbWhite", hexStr: "&hFFFFFF", expected: 16777215},
	}

	// 1. Verify in VBSConstants built-in catalog
	for _, ct := range colorTests {
		var found *VBSConstant
		for i := range VBSConstants {
			if VBSConstants[i].Name == ct.name {
				found = &VBSConstants[i]
				break
			}
		}
		if found == nil {
			t.Fatalf("Constant %s missing from VBSConstants catalog", ct.name)
		}
		if found.Val.Type != VTInteger {
			t.Fatalf("Constant %s is not VTInteger, got %v", ct.name, found.Val.Type)
		}
		if found.Val.Num != ct.expected {
			t.Errorf("VBSConstants[%q] = %d; expected %d (%s)", ct.name, found.Val.Num, ct.expected, ct.hexStr)
		}
	}

	// 2. Verify via VBScript runtime evaluation
	for _, ct := range colorTests {
		t.Run(ct.name, func(t *testing.T) {
			val := runVBScriptExpr(t, ct.name)
			if val.Type != VTInteger {
				t.Fatalf("expected VTInteger from %s, got type %v", ct.name, val.Type)
			}
			if val.Num != ct.expected {
				t.Errorf("%s evaluated to %d; expected %d (%s)", ct.name, val.Num, ct.expected, ct.hexStr)
			}
		})
	}
}
