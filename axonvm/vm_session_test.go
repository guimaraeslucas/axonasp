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
)

// TestVMSessionDispatch verifies native Session dispatch for values and counters.
func TestVMSessionDispatch(t *testing.T) {
	vm := NewVM(nil, nil, 5)
	host := NewMockHost()
	vm.SetHost(host)

	vm.dispatchNativeCall(3, "Set", []Value{NewString("Name"), NewString("Axon")})
	value := vm.dispatchNativeCall(3, "Get", []Value{NewString("name")})
	if value.Type != VTString || value.Str != "Axon" {
		t.Fatalf("expected string Axon, got %#v", value)
	}

	count := vm.dispatchNativeCall(3, "Count", nil)
	if count.Type != VTInteger || count.Num != 1 {
		t.Fatalf("expected count 1, got %#v", count)
	}

	exists := vm.dispatchNativeCall(3, "Exists", []Value{NewString("NAME")})
	if exists.Type != VTBool || exists.Num != 1 {
		t.Fatalf("expected Exists=True, got %#v", exists)
	}
}

// TestVMSessionProperties verifies Session property methods through VM dispatcher.
func TestVMSessionProperties(t *testing.T) {
	vm := NewVM(nil, nil, 5)
	host := NewMockHost()
	vm.SetHost(host)

	vm.dispatchNativeCall(3, "Timeout", []Value{NewInteger(40)})
	timeout := vm.dispatchNativeCall(3, "Timeout", nil)
	if timeout.Type != VTInteger || timeout.Num != 40 {
		t.Fatalf("expected Timeout 40, got %#v", timeout)
	}

	sessionID := vm.dispatchNativeCall(3, "SessionID", nil)
	if sessionID.Type != VTInteger || sessionID.Num < 0 {
		t.Fatalf("expected numeric SessionID, got %#v", sessionID)
	}
}

// TestVMSessionAbandon verifies abandon behavior through VM dispatcher.
func TestVMSessionAbandon(t *testing.T) {
	vm := NewVM(nil, nil, 5)
	host := NewMockHost()
	vm.SetHost(host)

	vm.dispatchNativeCall(3, "Set", []Value{NewString("k"), NewInteger(1)})
	vm.dispatchNativeCall(3, "Abandon", nil)

	value := vm.dispatchNativeCall(3, "Get", []Value{NewString("k")})
	if value.Type != VTEmpty {
		t.Fatalf("expected empty value after abandon, got %#v", value)
	}
}

// TestASPSessionContentsCollection verifies Session.Contents count/item/remove compatibility.
func TestASPSessionContentsCollection(t *testing.T) {
	source := `<%
Session("foo") = "bar"
Session("x") = "y"
Response.Write Session.Contents.Count
Response.Write "|"
Response.Write Session.Contents("foo")
Response.Write "|"
Session.Contents.Remove "foo"
Response.Write Session.Contents.Count
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
		t.Fatalf("expected Session.Contents output 2|bar|1, got %q", output.String())
	}
}

// TestASPSessionContentsEnumeration verifies For Each over Session.Contents yields keys.
func TestASPSessionContentsEnumeration(t *testing.T) {
	source := `<%
Session("k2") = "v2"
Session("k1") = "v1"
Dim key
For Each key In Session.Contents
	Response.Write key & ":" & Session.Contents(key) & "|"
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
	if rendered != "k1:v1|k2:v2|" {
		t.Fatalf("expected Session.Contents enumeration output, got %q", rendered)
	}
}

// TestASPSessionContentsKeys verifies Session.Contents.Keys() returns a VTArray-compatible ASP array.
func TestASPSessionContentsKeys(t *testing.T) {
	source := `<%
Session("k2") = "v2"
Session("k1") = "v1"
Dim keys
keys = Session.Contents.Keys()
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

	if output.String() != "k1|k2" {
		t.Fatalf("expected Session.Contents.Keys output k1|k2, got %q", output.String())
	}
}

// TestASPSessionStaticObjectsEnumeration verifies For Each over Session.StaticObjects returns only global.asa object keys.
func TestASPSessionStaticObjectsEnumeration(t *testing.T) {
	source := `<object runat="server" scope="Session" id="sessObjA" progid="Scripting.Dictionary"></object>
<object runat="server" scope="Session" id="sessObjB" progid="Scripting.FileSystemObject"></object>
<%
Dim key
For Each key In Session.StaticObjects
	Response.Write key & "|"
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
	if rendered != "\n\nsessobja|sessobjb|" {
		t.Fatalf("expected Session.StaticObjects keys output, got %q", rendered)
	}
}
