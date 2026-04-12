/*
 * AxonASP Server
 * Copyright (C) 2026 G3pix Ltda. All rights reserved.
 */
package axonvm

import (
	"bytes"
	"testing"
)

// TestASPConstForwardReference verifies global Const values resolve correctly even when read before declaration point.
func TestASPConstForwardReference(t *testing.T) {
	source := `<%
Response.Write JSON_ROOT_KEY
Const JSON_ROOT_KEY = "[[JSONroot]]"
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

	if output.String() != "[[JSONroot]]" {
		t.Fatalf("expected const output [[JSONroot]], got %q", output.String())
	}
}

// TestASPClassMethodUsesGlobalConst verifies class methods can resolve global constants safely.
func TestASPClassMethodUsesGlobalConst(t *testing.T) {
	source := `<%
Const JSON_ROOT_KEY = "[[JSONroot]]"
Const JSON_DEFAULT_PROPERTY_NAME = "data"

Class C
	Private Function F(ByVal prop)
		If prop = JSON_DEFAULT_PROPERTY_NAME Then
			F = JSON_ROOT_KEY
		Else
			F = prop
		End If
	End Function

	Public Function G()
		G = F("data")
	End Function
End Class

Dim o
Set o = New C
Response.Write o.G()
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

	if output.String() != "[[JSONroot]]" {
		t.Fatalf("expected class const output [[JSONroot]], got %q", output.String())
	}
}
