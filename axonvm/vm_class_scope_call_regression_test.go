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
	"errors"
	"testing"

	"g3pix.com.br/axonasp/v2/vbscript"
)

// TestVMClassMethodCallShadowsGlobal verifies that an unqualified class method
// call in statement position binds to the current class member even when a
// global variable with the same name exists.
func TestVMClassMethodCallShadowsGlobal(t *testing.T) {
	source := `<%
Dim Upload
Class U
	Public Value
	Private Sub Class_Initialize()
		Value = 0
	End Sub
	Public Sub Save()
		If Not False Then Upload
	End Sub
	Public Sub Upload()
		Value = 1
	End Sub
End Class
Set Upload = New U
Upload.Save
Response.Write Upload.Value
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

	if output.String() != "1" {
		t.Fatalf("expected class Upload method to be invoked, got output %q", output.String())
	}
}

// TestVMClassMethodCallsForwardDeclaredGlobalSub verifies that an implicit call
// from a Class method to a global Sub declared after the Class resolves to the
// global Sub and mutates ByRef arguments.
func TestVMClassMethodCallsForwardDeclaredGlobalSub(t *testing.T) {
	source := `<%
Class Example
	Function Run()
		Dim result
		LoadValue 1, result
		Run = result
	End Function
End Class

Sub LoadValue(id, result)
	result = True
End Sub

Dim instance
Set instance = New Example
Response.Write instance.Run()
%>`
	if got := runVBSAndGetOutput(t, source); got != "True" {
		t.Fatalf("expected True from forward global Sub call, got %q", got)
	}
}

// TestVMClassMethodCallsForwardDeclaredGlobalVariants verifies implicit calls,
// explicit Call statements, parenthesized single-arg calls, zero-arg Subs/Functions,
// case-insensitive resolution, and multiple classes sharing a forward global Sub.
func TestVMClassMethodCallsForwardDeclaredGlobalVariants(t *testing.T) {
	source := `<%
Class First
	Function RunImplicit()
		Dim result
		loadvalue 1, result
		RunImplicit = result
	End Function

	Function RunExplicitCall()
		Dim result
		Call LoadValue(1, result)
		RunExplicitCall = result
	End Function

	Function RunZeroArg()
		BumpCounter
		RunZeroArg = GetCounter
	End Function

	Function RunParenStmt()
		Dim x
		x = 0
		SetFlag(x)
		RunParenStmt = x & ":" & counter
	End Function
End Class

Class Second
	Function Check()
		Dim r
		LoadValue 2, r
		Check = r
	End Function
End Class

Dim counter
counter = 0

Sub LoadValue(id, result)
	result = (id > 0)
End Sub

Sub BumpCounter()
	counter = counter + 1
End Sub

Function GetCounter()
	GetCounter = counter
End Function

Sub SetFlag(ByRef flag)
	flag = 42
	counter = counter + 1
End Sub

Dim a, b
Set a = New First
Set b = New Second
Response.Write a.RunImplicit() & "|" & a.RunExplicitCall() & "|" & a.RunZeroArg() & "|" & a.RunParenStmt() & "|" & b.Check()
%>`
	want := "True|True|1|0:2|True"
	if got := runVBSAndGetOutput(t, source); got != want {
		t.Fatalf("unexpected output: got %q want %q", got, want)
	}
}

// TestVMClassMethodShadowsGlobalSub verifies that a Class method or property
// takes precedence over a global Sub with the same name regardless of order.
func TestVMClassMethodShadowsGlobalSub(t *testing.T) {
	source := `<%
Sub BeforeGlobal(ByRef outVal)
	outVal = "global-before"
End Sub

Class Shadowed
	Function CallMembers()
		Dim a, b
		BeforeGlobal a
		AfterGlobal b
		CallMembers = a & "|" & b
	End Function

	Sub BeforeGlobal(ByRef outVal)
		outVal = "class-before"
	End Sub

	Sub AfterGlobal(ByRef outVal)
		outVal = "class-after"
	End Sub
End Class

Sub AfterGlobal(ByRef outVal)
	outVal = "global-after"
End Sub

Dim obj
Set obj = New Shadowed
Response.Write obj.CallMembers()
%>`
	want := "class-before|class-after"
	if got := runVBSAndGetOutput(t, source); got != want {
		t.Fatalf("expected class members to shadow global subs, got %q want %q", got, want)
	}
}

// TestVMClassMissingMemberStillRaises01B6 verifies that genuine invalid
// property/method accesses on a class instance still raise 800A01B6.
func TestVMClassMissingMemberStillRaises01B6(t *testing.T) {
	cases := []string{
		`<%
Class Example
	Sub Run()
		Me.MissingMethod 1
	End Sub
End Class
Dim instance
Set instance = New Example
instance.Run
%>`,
		`<%
Class Example
	Sub Run()
		MissingMethod 1
	End Sub
End Class
Dim instance
Set instance = New Example
instance.Run
%>`,
		`<%
Class Example
End Class
Dim instance
Set instance = New Example
instance.MissingMethod
%>`,
	}

	for i, source := range cases {
		compiler := NewASPCompiler(source)
		if err := compiler.Compile(); err != nil {
			t.Fatalf("case %d compile failed: %v", i, err)
		}
		vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
		host := NewMockHost()
		vm.SetHost(host)
		err := vm.Run()
		if err == nil {
			t.Fatalf("case %d: expected 800A01B6 error, got nil", i)
		}
		var vmErr *VMError
		if !errors.As(err, &vmErr) {
			t.Fatalf("case %d: expected *VMError, got %T (%v)", i, err, err)
		}
		if vmErr.Code != vbscript.ObjectDoesntSupportThisPropertyOrMethod {
			t.Fatalf("case %d: expected code %d, got %d (%v)", i, vbscript.ObjectDoesntSupportThisPropertyOrMethod, vmErr.Code, vmErr)
		}
	}
}
