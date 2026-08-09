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
 */
package axonvm

import (
	"bytes"
	"testing"
)

// TestSynchronousClassTerminateSetNothing verifies that Class_Terminate runs
// synchronously on Set x = Nothing so that side-effects are observable on the
// immediate next line of VBScript.
func TestSynchronousClassTerminateSetNothing(t *testing.T) {
	source := `<%
Dim trace
trace = ""

Class TestObj
	Private Sub Class_Terminate()
		trace = trace & "[Terminated]"
	End Sub
End Class

Dim x
Set x = New TestObj
trace = trace & "[Created]"
Set x = Nothing
trace = trace & "[AfterSetNothing]"

Response.Write trace
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

	expected := "[Created][Terminated][AfterSetNothing]"
	if output.String() != expected {
		t.Fatalf("expected output %q, got %q", expected, output.String())
	}
}

// TestSynchronousClassTerminateLocalScopeExit verifies that when an object falls out
// of local scope in a Sub, Class_Terminate executes before the statement after the Sub call.
func TestSynchronousClassTerminateLocalScopeExit(t *testing.T) {
	source := `<%
Dim trace
trace = ""

Class ScopeObj
	Private Sub Class_Terminate()
		trace = trace & "[LocalTerminated]"
	End Sub
End Class

Sub DoWork()
	Dim tmp
	Set tmp = New ScopeObj
	trace = trace & "[InsideSub]"
End Sub

trace = trace & "[BeforeCall]"
DoWork
trace = trace & "[AfterCall]"

Response.Write trace
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

	expected := "[BeforeCall][InsideSub][LocalTerminated][AfterCall]"
	if output.String() != expected {
		t.Fatalf("expected output %q, got %q", expected, output.String())
	}
}

// TestSynchronousClassTerminateNestedDestruction verifies nested object destruction
// where Class_Terminate of parent object drops the reference to child object.
func TestSynchronousClassTerminateNestedDestruction(t *testing.T) {
	source := `<%
Dim trace
trace = ""

Class ChildObj
	Private Sub Class_Terminate()
		trace = trace & "[ChildTerminated]"
	End Sub
End Class

Class ParentObj
	Public Child
	Private Sub Class_Initialize()
		Set Child = New ChildObj
	End Sub
	Private Sub Class_Terminate()
		trace = trace & "[ParentTerminated]"
		Set Child = Nothing
	End Sub
End Class

Dim p
Set p = New ParentObj
trace = trace & "[Created]"
Set p = Nothing
trace = trace & "[AfterSetNothing]"

Response.Write trace
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

	expected := "[Created][ParentTerminated][ChildTerminated][AfterSetNothing]"
	if output.String() != expected {
		t.Fatalf("expected output %q, got %q", expected, output.String())
	}
}

// TestSynchronousClassTerminateCyclicProtection verifies that cyclic references do not cause
// infinite termination loops or crashes during Class_Terminate.
func TestSynchronousClassTerminateCyclicProtection(t *testing.T) {
	source := `<%
Dim trace
trace = ""

Class Node
	Public SelfRef
	Private Sub Class_Terminate()
		trace = trace & "[NodeTerminated]"
		Set SelfRef = Nothing
	End Sub
End Class

Dim n
Set n = New Node
Set n.SelfRef = n
trace = trace & "[Created]"
Set n = Nothing
trace = trace & "[Done]"

Response.Write trace
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

	expected := "[Created][NodeTerminated][Done]"
	if output.String() != expected {
		t.Fatalf("expected output %q, got %q", expected, output.String())
	}
}

// TestVMPoolCleanupAfterMultipleDestructors verifies that acquiring and releasing VMs
// with multiple destructors leaves the VM completely pristine for pool recycling.
func TestVMPoolCleanupAfterMultipleDestructors(t *testing.T) {
	source := `<%
Class Item
	Public Ref
	Private Sub Class_Terminate()
		' Clean up
	End Sub
End Class

Dim a, b, c
Set a = New Item
Set b = New Item
Set c = New Item
Set a.Ref = b
Set b.Ref = c

Set a = Nothing
Set b = Nothing
Set c = Nothing
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	cachedProg := buildCachedProgramFromCompiler(compiler)
	vm := AcquireVMFromCachedProgram(cachedProg)
	host := NewMockHost()
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}

	// Release returns the VM to pool after running resetForReuse
	vm.Release()

	// Re-acquire and verify state is completely clean
	vm2 := AcquireVMFromCachedProgram(cachedProg)
	if len(vm2.runtimeClassItems) != 0 {
		t.Fatalf("expected runtimeClassItems to be empty after pool release, got %d items", len(vm2.runtimeClassItems))
	}
	if len(vm2.classInstanceOrder) != 0 {
		t.Fatalf("expected classInstanceOrder to be empty, got %d items", len(vm2.classInstanceOrder))
	}
	if len(vm2.callStack) != 0 {
		t.Fatalf("expected callStack to be empty, got %d items", len(vm2.callStack))
	}
	if vm2.sp != -1 {
		t.Fatalf("expected sp = -1, got %d", vm2.sp)
	}
	vm2.Release()
}
