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

func TestVB6ClassUDTField(t *testing.T) {
	source := `<%
	Type Point
		X As Integer
		Y As Integer
	End Type

	Class PointHolder
		Private m_Pt As Point

		Public Sub SetPoint(x As Integer, y As Integer)
			m_Pt.X = x
			m_Pt.Y = y
		End Sub

		Public Function GetX() As Integer
			GetX = m_Pt.X
		End Function

		Public Function GetY() As Integer
			GetY = m_Pt.Y
		End Function
	End Class

	Dim holder
	Set holder = New PointHolder
	holder.SetPoint 42, 99

	Response.Write "X=" & holder.GetX() & "|Y=" & holder.GetY()
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
	expected := "X=42|Y=99"
	if out != expected {
		t.Fatalf("expected %q, got %q", expected, out)
	}
}

func TestVB6ClassUDTReturnAndParams(t *testing.T) {
	source := `<%
	Type Point
		X As Integer
		Y As Integer
	End Type

	Class PointFactory
		Public Function Create(x As Integer, y As Integer) As Point
			Create.X = x
			Create.Y = y
		End Function

		Public Function MovePoint(pt As Point, dx As Integer, dy As Integer) As Point
			MovePoint.X = pt.X + dx
			MovePoint.Y = pt.Y + dy
		End Function
	End Class

	Dim factory, p1, p2
	Set factory = New PointFactory
	p1 = factory.Create(10, 20)
	p2 = factory.MovePoint(p1, 5, -5)

	Response.Write "p1=" & p1.X & "," & p1.Y & "|p2=" & p2.X & "," & p2.Y
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
	expected := "p1=10,20|p2=15,15"
	if out != expected {
		t.Fatalf("expected %q, got %q", expected, out)
	}
}

func TestVB6RegressionNonObjectMemberSet(t *testing.T) {
	source := `<%
	Dim notAnObj
	notAnObj = 123
	On Error Resume Next
	notAnObj.X = 99
	Response.Write "Err=" & Err.Number & "|" & Err.Description
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
	// Object Required is HRESULT 424 (or 800A01A8/800A000D depending on context) in VBScript
	if !bytes.Contains(buf.Bytes(), []byte("required")) {
		t.Fatalf("expected 'Object required' error description, got: %q", out)
	}
}
