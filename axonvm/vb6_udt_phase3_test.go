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
	if !scanBytecodeForExtOp(bytecode, ExtOpInitRecord) {
		t.Fatal("expected ExtOpInitRecord in UDT bytecode")
	}
	if !scanBytecodeForExtOp(bytecode, ExtOpGetRecordMember) {
		t.Fatal("expected ExtOpGetRecordMember in UDT bytecode")
	}
	if !scanBytecodeForExtOp(bytecode, ExtOpSetRecordMember) {
		t.Fatal("expected ExtOpSetRecordMember in UDT bytecode")
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

func TestReleaseRecordHandlesSelfReference(t *testing.T) {
	vm := NewVM(nil, nil, 0)
	rec := vm.acquireRecord(1)
	rec.DefIdx = 1
	rec.Members[0] = Value{Type: VTRecord, Rec: rec}

	vm.releaseRecord(rec)
}

func TestVB6UDTPhase3RepeatedPooledExecution(t *testing.T) {
	source := `<%
Type Address
    City As String
End Type

Type Person
    Name As String
    Age As Integer
    Home As Address
End Type

Dim addr As Address
Dim p As Person

addr.City = "Sao Paulo"
p.Name = "Maya"
p.Age = 29
p.Home = addr

Response.Write p.Name & "|" & p.Age & "|" & p.Home.City
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	program := cachedProgramFromCompiler(compiler)
	for i := 0; i < 8; i++ {
		vm := AcquireVMFromCachedProgram(program)
		host := NewMockHost()
		var buf bytes.Buffer
		host.SetOutput(&buf)
		vm.SetHost(host)

		err := vm.Run()
		host.Response().Flush()
		got := buf.String()
		vm.Release()

		if err != nil {
			t.Fatalf("iteration %d: vm run failed: %v", i+1, err)
		}
		if got != "Maya|29|Sao Paulo" {
			t.Fatalf("iteration %d: unexpected output: %q", i+1, got)
		}
	}
}
