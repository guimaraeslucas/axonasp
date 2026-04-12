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
	"errors"
	"strconv"
	"testing"

	"g3pix.com.br/axonasp/vbscript"
)

// TestVMDimInitializerRejected verifies direct Dim initialization is rejected for Classic ASP compatibility.
func TestVMDimInitializerRejected(t *testing.T) {
	source := `<% Dim L = "Lucas" : Response.Write(L) %>`
	compiler := NewASPCompiler(source)
	err := compiler.Compile()
	if err == nil {
		t.Fatalf("expected compile failure for direct Dim initialization")
	}

	var syntaxErr *vbscript.VBSyntaxError
	if !errors.As(err, &syntaxErr) {
		t.Fatalf("expected VBScript syntax error, got: %T %v", err, err)
	}
	if syntaxErr.Code != vbscript.ExpectedEndOfStatement {
		t.Fatalf("unexpected syntax error code: got %d want %d", syntaxErr.Code, vbscript.ExpectedEndOfStatement)
	}
}

// TestVMDimAssignmentAfterDeclaration verifies Classic ASP style Dim followed by assignment still works.
func TestVMDimAssignmentAfterDeclaration(t *testing.T) {
	source := `<% Dim First : Dim Last : First = "Lu" : Last = "cas" : Response.Write(First & Last) %>`
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
		t.Fatalf("unexpected output: %q", output.String())
	}
}

// TestVMReDimBounds verifies Dim and ReDim array allocation works with VBScript bounds helpers.
func TestVMReDimBounds(t *testing.T) {
	source := `<% Dim Values() : ReDim Values(5) : Response.Write(LBound(Values)) : Response.Write("-") : Response.Write(UBound(Values)) %>`
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

	if output.String() != "0-5" {
		t.Fatalf("unexpected bounds output: %q", output.String())
	}
}

// TestVMReDimPreserveBounds verifies ReDim Preserve keeps array semantics while resizing.
func TestVMReDimPreserveBounds(t *testing.T) {
	source := `<% Dim Values(2) : ReDim Preserve Values(4) : Response.Write(TypeName(Values)) : Response.Write("|") : Response.Write(UBound(Values)) %>`
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

	if output.String() != "Variant()|4" {
		t.Fatalf("unexpected preserve output: %q", output.String())
	}
}

// TestVMArrayElementSetGet verifies array element assignment and indexed reads in ASP execution.
func TestVMArrayElementSetGet(t *testing.T) {
	source := `<% Dim items(2) : items(1) = "alpha" : items(2) = "beta" : Response.Write(items(1) & "," & items(2)) %>`
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

	if output.String() != "alpha,beta" {
		t.Fatalf("unexpected array indexed output: %q", output.String())
	}
}

// TestVMReDimPreserveMultiDimFirstDimensionChangeRejected verifies VBScript error 5 when Preserve changes non-last dimensions.
func TestVMReDimPreserveMultiDimFirstDimensionChangeRejected(t *testing.T) {
	source := `<%
On Error Resume Next
Dim arr()
ReDim arr(2, 3)
ReDim Preserve arr(3, 3)
Response.Write Err.Number
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

	expected := strconv.Itoa(vbscript.HRESULTFromVBScriptCode(vbscript.InvalidProcedureCallOrArgument))
	if output.String() != expected {
		t.Fatalf("unexpected Err.Number output: got %q want %q", output.String(), expected)
	}
}
