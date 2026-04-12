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
	"bytes"
	"testing"

	"g3pix.com.br/axonasp/axonvm/asp"
)

// TestVMApplicationDispatch verifies native Application member dispatch and value roundtrip.
func TestVMApplicationDispatch(t *testing.T) {
	vm := NewVM(nil, nil, 5)
	host := NewMockHost()
	vm.SetHost(host)

	vm.dispatchNativeCall(4, "Set", []Value{NewString("Counter"), NewInteger(42)})
	result := vm.dispatchNativeCall(4, "Get", []Value{NewString("counter")})

	if result.Type != VTInteger || result.Num != 42 {
		t.Fatalf("expected integer 42, got %#v", result)
	}

	count := vm.dispatchNativeCall(4, "Count", nil)
	if count.Type != VTInteger || count.Num != 1 {
		t.Fatalf("expected count 1, got %#v", count)
	}

	exists := vm.dispatchNativeCall(4, "Exists", []Value{NewString("COUNTER")})
	if exists.Type != VTBool || exists.Num != 1 {
		t.Fatalf("expected Exists=True, got %#v", exists)
	}

	vm.dispatchNativeCall(4, "Remove", []Value{NewString("counter")})
	result = vm.dispatchNativeCall(4, "Get", []Value{NewString("counter")})
	if result.Type != VTEmpty {
		t.Fatalf("expected VTEmpty after remove, got %#v", result)
	}
}

// TestVMApplicationLockDispatch verifies Lock and Unlock dispatch behavior.
func TestVMApplicationLockDispatch(t *testing.T) {
	vm := NewVM(nil, nil, 5)
	host := NewMockHost()
	vm.SetHost(host)

	vm.dispatchNativeCall(4, "Lock", nil)
	vm.dispatchNativeCall(4, "Lock", nil)

	lockCount := vm.dispatchNativeCall(4, "GetLockCount", nil)
	if lockCount.Type != VTInteger || lockCount.Num != 2 {
		t.Fatalf("expected lock count 2, got %#v", lockCount)
	}

	isLocked := vm.dispatchNativeCall(4, "IsLocked", nil)
	if isLocked.Type != VTBool || isLocked.Num != 1 {
		t.Fatalf("expected locked=True, got %#v", isLocked)
	}

	vm.dispatchNativeCall(4, "Unlock", nil)
	vm.dispatchNativeCall(4, "Unlock", nil)

	isLocked = vm.dispatchNativeCall(4, "IsLocked", nil)
	if isLocked.Type != VTBool || isLocked.Num != 0 {
		t.Fatalf("expected locked=False, got %#v", isLocked)
	}
}

// TestASPApplicationDefaultIndexer verifies Application("key") default member get/set semantics.
func TestASPApplicationDefaultIndexer(t *testing.T) {
	source := `<%
Application("Counter") = 10
Response.Write Application("Counter")
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

	if output.String() != "10" {
		t.Fatalf("expected Application default member output 10, got %q", output.String())
	}
}

// TestASPSessionDefaultIndexerWithArraySetPath verifies Session("key") assignment still works with indexed assignment opcode flow.
func TestASPSessionDefaultIndexerWithArraySetPath(t *testing.T) {
	source := `<%
Session("Name") = "Lucas"
Response.Write Session("Name")
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

	if output.String() != "Lucas" {
		t.Fatalf("expected Session default member output Lucas, got %q", output.String())
	}
}

// TestASPApplicationStaticObjectsEnumeration verifies For Each over Application.StaticObjects yields only global.asa object keys.
func TestASPApplicationStaticObjectsEnumeration(t *testing.T) {
	source := `<%
Dim key
For Each key In Application.StaticObjects
	Response.Write key & "|"
Next
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	host.Application().AddStaticObject("item2", asp.NewApplicationString(staticObjectProgIDPrefix+"Scripting.Dictionary"))
	host.Application().AddStaticObject("item1", asp.NewApplicationString(staticObjectProgIDPrefix+"Scripting.FileSystemObject"))
	host.Application().AddStaticObject("nonObject", asp.NewApplicationString("plainValue"))
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	rendered := output.String()
	if rendered != "item1|item2|" {
		t.Fatalf("expected static object keys output, got %q", rendered)
	}
}

// TestASPApplicationContentsCollection verifies Application.Contents count/item/remove compatibility.
func TestASPApplicationContentsCollection(t *testing.T) {
	source := `<%
Application("foo") = "bar"
Application("x") = "y"
Response.Write Application.Contents.Count
Response.Write "|"
Response.Write Application.Contents("foo")
Response.Write "|"
Application.Contents.Remove "foo"
Response.Write Application.Contents.Count
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

	if output.String() != "2|bar|1" {
		t.Fatalf("expected Application.Contents output 2|bar|1, got %q", output.String())
	}
}

// TestASPApplicationContentsEnumeration verifies For Each over Application.Contents yields keys.
func TestASPApplicationContentsEnumeration(t *testing.T) {
	source := `<%
Application("b") = "2"
Application("a") = "1"
Dim key
For Each key In Application.Contents
	Response.Write key & ":" & Application.Contents(key) & "|"
Next
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

	rendered := output.String()
	if rendered != "a:1|b:2|" {
		t.Fatalf("expected Application.Contents enumeration output, got %q", rendered)
	}
}

// TestASPApplicationContentsKeys verifies Application.Contents.Keys() returns a VTArray-compatible ASP array.
func TestASPApplicationContentsKeys(t *testing.T) {
	source := `<%
Application("b") = "2"
Application("a") = "1"
Dim keys
keys = Application.Contents.Keys()
If IsArray(keys) Then
	Response.Write Join(keys, "|")
Else
	Response.Write "NOT_ARRAY"
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

	if output.String() != "a|b" {
		t.Fatalf("expected Application.Contents.Keys output a|b, got %q", output.String())
	}
}

// TestASPApplicationStaticObjectsCountProperty verifies Application.StaticObjects.Count compatibility.
func TestASPApplicationStaticObjectsCountProperty(t *testing.T) {
	source := `<%= Application.StaticObjects.Count %>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	host.Application().AddStaticObject("x", asp.NewApplicationString("1"))
	host.Application().AddStaticObject("y", asp.NewApplicationString("2"))
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}

	if output.String() != "2" {
		t.Fatalf("expected StaticObjects.Count = 2, got %q", output.String())
	}
}
