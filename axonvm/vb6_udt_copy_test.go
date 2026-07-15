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

func TestVB6UDTCopySemantics(t *testing.T) {
	source := `<%
	Type Point
		X As Integer
		Y As Integer
	End Type

	Dim p As Point, p2 As Point
	p.X = 10
	p.Y = 20
	p2 = p
	p2.X = 99

	Response.Write "p.X=" & p.X & "|p2.X=" & p2.X
	%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
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
	expected := "p.X=10|p2.X=99"
	if out != expected {
		t.Fatalf("expected %q, got %q", expected, out)
	}
}

func TestVB6NestedUDTCopySemantics(t *testing.T) {
	source := `<%
	Type Address
		City As String
		Zip As Integer
	End Type

	Type User
		Name As String
		Home As Address
	End Type

	Dim u As User
	Dim a As Address
	u.Name = "G3pix"
	a.City = "Floripa"
	a.Zip = 88000
	u.Home = a

	' Mutate original address after assignment
	a.City = "Porto"

	Response.Write "u.Home.City=" & u.Home.City & "|a.City=" & a.City
	%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
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
	expected := "u.Home.City=Floripa|a.City=Porto"
	if out != expected {
		t.Fatalf("expected %q, got %q", expected, out)
	}
}

func TestVB6UDTArrayCopySemantics(t *testing.T) {
	source := `<%
	Type Point
		X As Integer
		Y As Integer
	End Type

	Dim pts(1) As Point
	Dim p0 As Point
	p0.X = 10
	p0.Y = 20
	pts(0) = p0

	' Mutate original point after array assignment
	p0.X = 99

	Response.Write "pts(0).X=" & pts(0).X & "|p0.X=" & p0.X
	%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
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
	expected := "pts(0).X=10|p0.X=99"
	if out != expected {
		t.Fatalf("expected %q, got %q", expected, out)
	}
}
