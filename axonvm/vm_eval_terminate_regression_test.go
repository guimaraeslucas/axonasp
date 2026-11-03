//go:build !wasm

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
	"strings"
	"testing"
)

// runEvalTerminateASP compiles and runs one ASP page through the compiler + VM pipeline
// and returns the rendered output together with the run error.
func runEvalTerminateASP(t *testing.T, source string) (string, error) {
	t.Helper()

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	runErr := vm.Run()
	host.Response().Flush()

	return output.String(), runErr
}

// assertEvalTerminateClean fails the test when the page raised any runtime error or
// produced a payload different from the expected exact output.
func assertEvalTerminateClean(t *testing.T, source, expected string) {
	t.Helper()

	output, err := runEvalTerminateASP(t, source)
	if err != nil {
		t.Fatalf("unexpected runtime error: %v", err)
	}
	if strings.Contains(output, "800A01A8") {
		t.Fatalf("output leaked the 800A01A8 error payload: %q", output)
	}
	if output != expected {
		t.Fatalf("expected output %q, got %q", expected, output)
	}
}

// TestEvalTerminateEvalCallInFunction covers the minimal dynamic Eval(s) case inside a
// Function while one page-level instance stays alive until the end of the request.
func TestEvalTerminateEvalCallInFunction(t *testing.T) {
	source := `<%
Class Item
  Private Sub Class_Terminate()
    Response.Write "[term]"
  End Sub
End Class
Set keep = New Item

Class Ctl
  Public Sub Go()
  End Sub
End Class

Function Route()
  Dim ci, s
  Set ci = New Ctl
  s = "ci.Go()"
  Call Eval(s)
End Function

Route
Response.Write "done"
%>`

	assertEvalTerminateClean(t, source, "done[term]")
}

// TestEvalTerminateLiteralEval covers Call Eval("ci.Go()") with a literal string argument.
func TestEvalTerminateLiteralEval(t *testing.T) {
	source := `<%
Class Item
  Private Sub Class_Terminate()
    Response.Write "[term]"
  End Sub
End Class
Set keep = New Item

Class Ctl
  Public Sub Go()
  End Sub
End Class

Function Route()
  Dim ci, s
  Set ci = New Ctl
  Call Eval("ci.Go()")
End Function

Route
Response.Write "done"
%>`

	assertEvalTerminateClean(t, source, "done[term]")
}

// TestEvalTerminateRouterChained covers Eval("new " + "Ctl"), Eval(s) and the chained
// f.Find().Value expression that materializes one temporary instance inside the Eval scope.
func TestEvalTerminateRouterChained(t *testing.T) {
	source := `<%
Class Item
  Private Sub Class_Terminate()
    Response.Write "[term]"
  End Sub
  Public Property Get Value()
    Value = "x"
  End Property
End Class
Set keep = New Item

Class Finder
  Public Function Find()
    Set Find = New Item
  End Function
End Class

Class Ctl
  Public Sub Go()
    Set f = New Finder
    v = f.Find().Value
  End Sub
End Class

Function Route()
  Dim ci, s
  Set ci = Eval("new " + "Ctl")
  s = "ci.Go()"
  Call Eval(s)
End Function

Route
Response.Write "done"
%>`

	output, err := runEvalTerminateASP(t, source)
	if err != nil {
		t.Fatalf("unexpected runtime error: %v", err)
	}
	// The temporary Finder.Find() instance is never bound to a slot, so it is swept at the
	// end of the request together with the page-level keep instance. Both Class_Terminate
	// bodies must run in reverse construction order after "done".
	if output != "done[term][term]" {
		t.Fatalf("expected %q, got %q", "done[term][term]", output)
	}
}

// TestEvalTerminateWorkaroundIntermediate covers the assigned-intermediate variant
// Set it = f.Find() followed by v = it.Value.
func TestEvalTerminateWorkaroundIntermediate(t *testing.T) {
	source := `<%
Class Item
  Private Sub Class_Terminate()
    Response.Write "[term]"
  End Sub
  Public Property Get Value()
    Value = "x"
  End Property
End Class
Set keep = New Item

Class Finder
  Public Function Find()
    Set Find = New Item
  End Function
End Class

Class Ctl
  Public Sub Go()
    Set f = New Finder
    Set it = f.Find()
    v = it.Value
  End Sub
End Class

Function Route()
  Dim ci, s
  Set ci = Eval("new " + "Ctl")
  s = "ci.Go()"
  Call Eval(s)
End Function

Route
Response.Write "done"
%>`

	output, err := runEvalTerminateASP(t, source)
	if err != nil {
		t.Fatalf("unexpected runtime error: %v", err)
	}
	// The intermediate instance bound to the local "it" is released when Go returns, so its
	// destructor runs inside the Eval scope before "done", and keep terminates afterwards.
	if output != "[term]done[term]" {
		t.Fatalf("expected %q, got %q", "[term]done[term]", output)
	}
}

// TestEvalTerminateTwoObjectsKept covers two instances instantiated inside an Eval-invoked
// action and kept in page-level slots until the end of the request.
func TestEvalTerminateTwoObjectsKept(t *testing.T) {
	source := `<%
Class Item
  Public Name
  Private Sub Class_Terminate()
    Response.Write "[term:" & Name & "]"
  End Sub
End Class
Dim kept1, kept2

Class Ctl
  Public Sub Go()
    Set kept1 = New Item : kept1.Name = "1"
    Set kept2 = New Item : kept2.Name = "2"
  End Sub
End Class

Function Route()
  Dim ci, s
  Set ci = Eval("new " + "Ctl")
  s = "ci.Go()"
  Call Eval(s)
End Function

Route
Response.Write "done"
%>`

	assertEvalTerminateClean(t, source, "done[term:2][term:1]")
}

// TestEvalTerminateOneObjectKept covers a single instance instantiated inside an
// Eval-invoked action and kept in a page-level slot until the end of the request.
func TestEvalTerminateOneObjectKept(t *testing.T) {
	source := `<%
Class Item
  Public Name
  Private Sub Class_Terminate()
    Response.Write "[term:" & Name & "]"
  End Sub
End Class
Dim kept1

Class Ctl
  Public Sub Go()
    Set kept1 = New Item : kept1.Name = "1"
  End Sub
End Class

Function Route()
  Dim ci, s
  Set ci = Eval("new " + "Ctl")
  s = "ci.Go()"
  Call Eval(s)
End Function

Route
Response.Write "done"
%>`

	assertEvalTerminateClean(t, source, "done[term:1]")
}

// TestEvalTerminateReleasedBeforeEnd covers two instances created through Eval that are
// explicitly cleared before the page ends, so both terminate before "done" is written.
func TestEvalTerminateReleasedBeforeEnd(t *testing.T) {
	source := `<%
Class Item
  Public Name
  Private Sub Class_Terminate()
    Response.Write "[term:" & Name & "]"
  End Sub
End Class
Dim kept1, kept2

Class Ctl
  Public Sub Go()
    Set kept1 = New Item : kept1.Name = "1"
    Set kept2 = New Item : kept2.Name = "2"
  End Sub
End Class

Function Route()
  Dim ci, s
  Set ci = Eval("new " + "Ctl")
  s = "ci.Go()"
  Call Eval(s)
End Function

Route
Set kept1 = Nothing
Set kept2 = Nothing
Response.Write "done"
%>`

	assertEvalTerminateClean(t, source, "[term:1][term:2]done")
}

// TestEvalTerminateNoEvalDirect is the control scenario: the controller is created and
// invoked directly, without any dynamic Eval, and must behave identically to the Eval path.
func TestEvalTerminateNoEvalDirect(t *testing.T) {
	source := `<%
Class Item
  Public Name
  Private Sub Class_Terminate()
    Response.Write "[term:" & Name & "]"
  End Sub
End Class
Dim kept1, kept2

Class Ctl
  Public Sub Go()
    Set kept1 = New Item : kept1.Name = "1"
    Set kept2 = New Item : kept2.Name = "2"
  End Sub
End Class

Dim ci
Set ci = New Ctl
ci.Go
Response.Write "done"
%>`

	assertEvalTerminateClean(t, source, "done[term:2][term:1]")
}
