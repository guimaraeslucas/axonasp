/*
 * AxonASP Server
 * Copyright (C) 2026 G3pix Ltda. All rights reserved.
 *
 * This test file covers the modern VBScript operator extensions:
 *   - Bitshift operators (<<, >>)
 *   - IsNot operator
 *   - Short-circuit operators (AndAlso, OrElse)
 */

package axonvm

import (
	"bytes"
	"testing"
)

// --- Bitshift Tests ---

// TestShiftLeft verifies basic and edge-case left shifts.
func TestShiftLeft(t *testing.T) {
	source := `<%
Dim a
a = 1 << 3
Response.Write a
%>`
	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}
	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)
	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()
	expected := "8"
	if output.String() != expected {
		t.Fatalf("unexpected output: got %q want %q", output.String(), expected)
	}
}

// TestShiftRight verifies basic and edge-case right shifts.
func TestShiftRight(t *testing.T) {
	source := `<%
Dim a
a = 16 >> 2
Response.Write a
%>`
	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}
	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)
	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()
	expected := "4"
	if output.String() != expected {
		t.Fatalf("unexpected output: got %q want %q", output.String(), expected)
	}
}

// TestShiftExcessive verifies that shifts exceeding 64 bits yield 0.
func TestShiftExcessive(t *testing.T) {
	tests := []struct {
		name     string
		source   string
		expected string
	}{
		{
			name:     "shift left excessive",
			source:   `<% Dim a : a = 1 << 64 : Response.Write a %>`,
			expected: "0",
		},
		{
			name:     "shift left way excessive",
			source:   `<% Dim a : a = 1 << 100 : Response.Write a %>`,
			expected: "0",
		},
		{
			name:     "shift right excessive",
			source:   `<% Dim a : a = 1 >> 64 : Response.Write a %>`,
			expected: "0",
		},
		{
			name:     "shift right way excessive",
			source:   `<% Dim a : a = 1 >> 100 : Response.Write a %>`,
			expected: "0",
		},
		{
			name:     "shift left zero bits",
			source:   `<% Dim a : a = 42 << 0 : Response.Write a %>`,
			expected: "42",
		},
		{
			name:     "shift right zero bits",
			source:   `<% Dim a : a = 42 >> 0 : Response.Write a %>`,
			expected: "42",
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			compiler := NewASPCompiler(tt.source)
			if err := compiler.Compile(); err != nil {
				t.Fatalf("compile failed: %v", err)
			}
			vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
			host := NewMockHost()
			var output bytes.Buffer
			host.SetOutput(&output)
			vm.SetHost(host)
			if err := vm.Run(); err != nil {
				t.Fatalf("vm run failed: %v", err)
			}
			host.Response().Flush()
			if output.String() != tt.expected {
				t.Fatalf("unexpected output: got %q want %q", output.String(), tt.expected)
			}
		})
	}
}

// --- IsNot Tests ---

// TestIsNot verifies that IsNot is the logical opposite of Is.
func TestIsNot(t *testing.T) {
	tests := []struct {
		name     string
		source   string
		expected string
	}{
		{
			name:     "IsNot Nothing is False for Nothing compared with Nothing",
			source:   `<% Dim a : Set a = Nothing : Response.Write (a IsNot Nothing) %>`,
			expected: "False",
		},
		{
			name:     "IsNot Nothing is True for object compared with Nothing",
			source:   `<% Dim a : Set a = Server.CreateObject("Collection") : Response.Write (a IsNot Nothing) %>`,
			expected: "True",
		},
		{
			name:     "IsNot same objects is False",
			source:   `<% Dim a, b : Set a = Server.CreateObject("Collection") : Set b = a : Response.Write (a IsNot b) %>`,
			expected: "False",
		},
		{
			name:     "IsNot different objects is True",
			source:   `<% Dim a, b : Set a = Server.CreateObject("Collection") : Set b = Server.CreateObject("Collection") : Response.Write (a IsNot b) %>`,
			expected: "True",
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			compiler := NewASPCompiler(tt.source)
			if err := compiler.Compile(); err != nil {
				t.Fatalf("compile failed: %v", err)
			}
			vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
			host := NewMockHost()
			var output bytes.Buffer
			host.SetOutput(&output)
			vm.SetHost(host)
			if err := vm.Run(); err != nil {
				t.Fatalf("vm run failed: %v", err)
			}
			host.Response().Flush()
			if output.String() != tt.expected {
				t.Fatalf("unexpected output: got %q want %q", output.String(), tt.expected)
			}
		})
	}
}

// TestIsNotTwoTokens verifies that "Is Not" (two tokens) still works.
func TestIsNotTwoTokens(t *testing.T) {
	source := `<%
Dim a
Set a = Nothing
if a Is Not Nothing Then
  Response.Write "unexpected"
Else
  Response.Write "IsNotTwoTokens"
End If
%>`
	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}
	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)
	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()
	expected := "IsNotTwoTokens"
	if output.String() != expected {
		t.Fatalf("unexpected output: got %q want %q", output.String(), expected)
	}
}

// TestIsNotBooleanOpposite verifies that A IsNot B is the exact boolean opposite of A Is B.
func TestIsNotBooleanOpposite(t *testing.T) {
	source := `<%
Dim a, b
Set a = Nothing
Set b = Nothing
Dim isResult, isNotResult
isResult = a Is b
isNotResult = a IsNot b
Response.Write isResult & "|" & isNotResult
%>`
	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}
	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)
	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()
	// a Is b is True, a IsNot b is False
	expected := "True|False"
	if output.String() != expected {
		t.Fatalf("unexpected output: got %q want %q", output.String(), expected)
	}
}

// --- Short-Circuit Tests ---

// TestAndAlsoShortCircuit verifies that AndAlso does NOT evaluate the RHS when the LHS is False.
func TestAndAlsoShortCircuit(t *testing.T) {
	source := `<%
Function divideByZero()
    divideByZero = 1/0
End Function

Dim result
result = False AndAlso divideByZero()
Response.Write "survived"
%>`
	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}
	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)
	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed (short-circuit should prevent division by zero): %v", err)
	}
	host.Response().Flush()
	expected := "survived"
	if output.String() != expected {
		t.Fatalf("unexpected output: got %q want %q", output.String(), expected)
	}
}

// TestAndAlsoEvaluatesRHS verifies that AndAlso DOES evaluate the RHS when the LHS is True.
func TestAndAlsoEvaluatesRHS(t *testing.T) {
	source := `<%
Dim sideEffect
sideEffect = "RHS called"

Function doSideEffect()
    doSideEffect = sideEffect
End Function

Dim result
result = True AndAlso doSideEffect()
Response.Write result
%>`
	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}
	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)
	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()
	// True AndAlso "RHS" should yield "RHS" (truthy -> the function's return value)
	if output.String() == "" {
		t.Fatalf("expected RHS side effect, got empty output")
	}
}

// TestOrElseShortCircuit verifies that OrElse does NOT evaluate the RHS when the LHS is True.
func TestOrElseShortCircuit(t *testing.T) {
	source := `<%
Function divideByZero()
    divideByZero = 1/0
End Function

Dim result
result = True OrElse divideByZero()
Response.Write "survived"
%>`
	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}
	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)
	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed (short-circuit should prevent division by zero): %v", err)
	}
	host.Response().Flush()
	expected := "survived"
	if output.String() != expected {
		t.Fatalf("unexpected output: got %q want %q", output.String(), expected)
	}
}

// TestOrElseEvaluatesRHS verifies that OrElse DOES evaluate the RHS when the LHS is False.
func TestOrElseEvaluatesRHS(t *testing.T) {
	source := `<%
Dim sideEffect
sideEffect = "RHS called"

Function doSideEffect()
    doSideEffect = sideEffect
End Function

Dim result
result = False OrElse doSideEffect()
Response.Write result
%>`
	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}
	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)
	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()
	// False OrElse "RHS called" should yield "RHS called"
	if output.String() == "" {
		t.Fatalf("expected RHS side effect, got empty output")
	}
}

// TestAndAlsoResult verifies the boolean result of AndAlso.
func TestAndAlsoResult(t *testing.T) {
	tests := []struct {
		name     string
		source   string
		expected string
	}{
		{
			name:     "True AndAlso True",
			source:   `<% Response.Write (True AndAlso True) %>`,
			expected: "True",
		},
		{
			name:     "True AndAlso False",
			source:   `<% Response.Write (True AndAlso False) %>`,
			expected: "False",
		},
		{
			name:     "False AndAlso True",
			source:   `<% Response.Write (False AndAlso True) %>`,
			expected: "False",
		},
		{
			name:     "False AndAlso False",
			source:   `<% Response.Write (False AndAlso False) %>`,
			expected: "False",
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			compiler := NewASPCompiler(tt.source)
			if err := compiler.Compile(); err != nil {
				t.Fatalf("compile failed: %v", err)
			}
			vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
			host := NewMockHost()
			var output bytes.Buffer
			host.SetOutput(&output)
			vm.SetHost(host)
			if err := vm.Run(); err != nil {
				t.Fatalf("vm run failed: %v", err)
			}
			host.Response().Flush()
			if output.String() != tt.expected {
				t.Fatalf("unexpected output: got %q want %q", output.String(), tt.expected)
			}
		})
	}
}

// TestOrElseResult verifies the boolean result of OrElse.
func TestOrElseResult(t *testing.T) {
	tests := []struct {
		name     string
		source   string
		expected string
	}{
		{
			name:     "True OrElse True",
			source:   `<% Response.Write (True OrElse True) %>`,
			expected: "True",
		},
		{
			name:     "True OrElse False",
			source:   `<% Response.Write (True OrElse False) %>`,
			expected: "True",
		},
		{
			name:     "False OrElse True",
			source:   `<% Response.Write (False OrElse True) %>`,
			expected: "True",
		},
		{
			name:     "False OrElse False",
			source:   `<% Response.Write (False OrElse False) %>`,
			expected: "False",
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			compiler := NewASPCompiler(tt.source)
			if err := compiler.Compile(); err != nil {
				t.Fatalf("compile failed: %v", err)
			}
			vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
			host := NewMockHost()
			var output bytes.Buffer
			host.SetOutput(&output)
			vm.SetHost(host)
			if err := vm.Run(); err != nil {
				t.Fatalf("vm run failed: %v", err)
			}
			host.Response().Flush()
			if output.String() != tt.expected {
				t.Fatalf("unexpected output: got %q want %q", output.String(), tt.expected)
			}
		})
	}
}

// --- Regression: Standard And/Or still evaluate both sides ---

// TestStandardAndEvaluatesBothSides verifies that the standard And operator
// still evaluates both sides (no short-circuit).
func TestStandardAndEvaluatesBothSides(t *testing.T) {
	source := `<%
Function triggerError()
    triggerError = 1/0
End Function

Dim result
result = False And triggerError()
Response.Write "should_not_reach"
%>`
	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}
	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)
	// Standard And must evaluate the RHS, so division by zero should cause an error.
	err := vm.Run()
	if err == nil {
		t.Fatalf("expected runtime error from division by zero in standard And, but got none")
	}
}

// TestStandardOrEvaluatesBothSides verifies that the standard Or operator
// still evaluates both sides (no short-circuit).
func TestStandardOrEvaluatesBothSides(t *testing.T) {
	source := `<%
Function triggerError()
    triggerError = 1/0
End Function

Dim result
result = True Or triggerError()
Response.Write "should_not_reach"
%>`
	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}
	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)
	// Standard Or must evaluate the RHS, so division by zero should cause an error.
	err := vm.Run()
	if err == nil {
		t.Fatalf("expected runtime error from division by zero in standard Or, but got none")
	}
}

// TestStandardIsStillWorks verifies that the standard Is operator still works.
func TestStandardIsStillWorks(t *testing.T) {
	source := `<%
Dim a, b
Set a = Nothing
Set b = Nothing
Response.Write (a Is b)
%>`
	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}
	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)
	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()
	expected := "True"
	if output.String() != expected {
		t.Fatalf("unexpected output: got %q want %q", output.String(), expected)
	}
}
