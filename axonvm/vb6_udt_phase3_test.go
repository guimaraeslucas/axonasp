/*
 * AxonASP Server
 * Copyright (C) 2026 G3pix Ltda. All rights reserved.
 *
 * Developed by Lucas Guimaraes - G3pix Ltda
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

func TestVB6UDTPhase3CompileAndRun(t *testing.T) {
	source := `<%
Type Address
    City As String
End Type

Type Person
    Name As String
    Age As Integer
    Home As Address
End Type

Dim a As Address
Dim p As Person

a.City = "Porto"
p.Name = "Lia"
p.Age = 33
p.Home = a

Response.Write p.Name & "|" & p.Age & "|" & p.Home.City
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	bytecode := compiler.Bytecode()
	if !scanBytecodeForOp(bytecode, OpInitRecord) {
		t.Fatal("expected OpInitRecord in UDT bytecode")
	}
	if !scanBytecodeForOp(bytecode, OpGetRecordMember) {
		t.Fatal("expected OpGetRecordMember in UDT bytecode")
	}
	if !scanBytecodeForOp(bytecode, OpSetRecordMember) {
		t.Fatal("expected OpSetRecordMember in UDT bytecode")
	}

	vm := NewVMFromCompiler(compiler)
	host := NewMockHost()
	var buf bytes.Buffer
	host.SetOutput(&buf)
	vm.SetHost(host)
	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()
	out := buf.String()
	if out != "Lia|33|Porto" {
		t.Fatalf("unexpected output: %q", out)
	}
}

func TestVB6UDTRetrocompatVariantPath(t *testing.T) {
	source := `<%
Dim v
v = "legacy"
v = v & "-asp"
Response.Write v
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	out := runVBSAndGetOutput(t, source)
	if out != "legacy-asp" {
		t.Fatalf("unexpected output: %q", out)
	}
}
