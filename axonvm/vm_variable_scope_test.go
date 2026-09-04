package axonvm

import (
	"bytes"
	"strings"
	"testing"
)

// TestImplicitPageScopeVariableResolution verifies that implicitly created page-scope variables
// (without Dim) resolve correctly inside Function and Sub bodies.
func TestImplicitPageScopeVariableResolution(t *testing.T) {
	source := `<%
A = "page-set"

Function ReadInFunction()
	ReadInFunction = "[" & A & "]"
End Function

Sub ReadInSub()
	Response.Write "  read inside sub      : [" & A & "]" & vbCrLf
End Sub

Sub WriteInSub()
	A = "sub-set"
End Sub

Dim D
D = "page-set"
Function ReadDimInFunction()
	ReadDimInFunction = "[" & D & "]"
End Function

Response.Write "  read at page scope   : [" & A & "]" & vbCrLf
Response.Write "  read inside function : " & ReadInFunction() & vbCrLf
ReadInSub()
WriteInSub()
Response.Write "  after sub assigned   : [" & A & "]" & vbCrLf
Response.Write "  Dim'd control        : " & ReadDimInFunction() & vbCrLf
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVMFromCompiler(compiler)
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	got := output.String()

	expectedSubstrings := []struct {
		desc string
		want string
	}{
		{desc: "read at page scope", want: "read at page scope   : [page-set]"},
		{desc: "read inside function", want: "read inside function : [page-set]"},
		{desc: "read inside sub", want: "read inside sub      : [page-set]"},
		{desc: "after sub assigned", want: "after sub assigned   : [sub-set]"},
		{desc: "Dim'd control", want: "Dim'd control        : [page-set]"},
	}

	for _, tc := range expectedSubstrings {
		if !strings.Contains(got, tc.want) {
			t.Errorf("Assertion failure for %s:\nwant substring: %q\ngot output:\n%s", tc.desc, tc.want, got)
		}
	}
}

// TestLocalDimShadowsPageScopeVariable verifies that explicitly Dim'd variables inside a
// procedure shadow page-scope variables and do not mutate the page-scope instance.
func TestLocalDimShadowsPageScopeVariable(t *testing.T) {
	source := `<%
A = "page-val"

Sub LocalShadowSub()
	Dim A
	A = "local-val"
	Response.Write "local: [" & A & "]" & vbCrLf
End Sub

LocalShadowSub()
Response.Write "page: [" & A & "]" & vbCrLf
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVMFromCompiler(compiler)
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	got := output.String()
	if !strings.Contains(got, "local: [local-val]") {
		t.Errorf("expected local-val inside sub, got:\n%s", got)
	}
	if !strings.Contains(got, "page: [page-val]") {
		t.Errorf("expected page-val at page scope, got:\n%s", got)
	}
}

// TestUndeclaredVariableInSubDoesNotLeakToPageScope verifies that a variable assigned
// solely inside a procedure without existing at page scope remains local to the procedure.
func TestUndeclaredVariableInSubDoesNotLeakToPageScope(t *testing.T) {
	source := `<%
Sub LocalOnlySub()
	IsolatedLocal = "sub-val"
	Response.Write "inside: [" & IsolatedLocal & "]" & vbCrLf
End Sub

LocalOnlySub()
Response.Write "outside: [" & IsolatedLocal & "]" & vbCrLf
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVMFromCompiler(compiler)
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	got := output.String()
	if !strings.Contains(got, "inside: [sub-val]") {
		t.Errorf("expected sub-val inside sub, got:\n%s", got)
	}
	if !strings.Contains(got, "outside: []") {
		t.Errorf("expected empty at page scope, got:\n%s", got)
	}
}

// TestImplicitPageScopeVariableAssignedAfterProcedure verifies that a page-scope variable
// assigned after procedure declarations still resolves to the page-scope variable.
func TestImplicitPageScopeVariableAssignedAfterProcedure(t *testing.T) {
	source := `<%
Function ReadLateAssigned()
	ReadLateAssigned = "[" & LateVar & "]"
End Function

LateVar = "assigned-late"

Response.Write "result: " & ReadLateAssigned() & vbCrLf
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVMFromCompiler(compiler)
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	got := output.String()
	if !strings.Contains(got, "result: [assigned-late]") {
		t.Errorf("expected [assigned-late], got:\n%s", got)
	}
}

// TestOptionExplicitRejectsImplicitPageScopeVariable verifies that Option Explicit continues
// to enforce variable declaration and rejects implicitly created page-scope variables.
func TestOptionExplicitRejectsImplicitPageScopeVariable(t *testing.T) {
	source := `<%
Option Explicit
A = "page-set"
%>`

	compiler := NewASPCompiler(source)
	err := compiler.Compile()
	if err == nil {
		t.Fatalf("expected compile error under Option Explicit for undeclared variable 'A', got nil")
	}
	if !strings.Contains(err.Error(), "Variable not defined") {
		t.Errorf("expected 'Variable not defined' error, got: %v", err)
	}
}
