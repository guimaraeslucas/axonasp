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
	"strings"
	"testing"
)

func TestVB6Interfaces(t *testing.T) {
	code := `
		Class IAnimal
			Function MakeSound()
			End Function
		End Class

		Class Dog
			Implements IAnimal
			
			Function IAnimal_MakeSound()
				IAnimal_MakeSound = "Woof!"
			End Function
			
			Function MakeSound()
				MakeSound = "Generic Dog Sound"
			End Function
		End Class

		Dim obj As IAnimal
		Set obj = New Dog
		Response.Write "Typed: " & obj.MakeSound()
		
		Dim obj2
		Set obj2 = New Dog
		Response.Write " | Untyped: " & obj2.MakeSound()
	`

	output, err := runVBScriptTest(code)
	if err != nil {
		t.Fatalf("Execution failed: %v", err)
	}

	expected := "Typed: Woof! | Untyped: Generic Dog Sound"
	if !strings.Contains(output, expected) {
		t.Errorf("Expected output to contain %q, got %q", expected, output)
	}
}

func TestVB6InterfaceProperty(t *testing.T) {
	code := `
		Class IData
			Property Get Value()
			End Property
		End Class

		Class MyData
			Implements IData
			Private m_val
			Sub Class_Initialize()
				m_val = "Hidden"
			End Sub
			Property Get IData_Value()
				IData_Value = m_val
			End Property
		End Class

		Dim d As IData
		Set d = New MyData
		Response.Write d.Value
	`

	output, err := runVBScriptTest(code)
	if err != nil {
		t.Fatalf("Execution failed: %v", err)
	}

	expected := "Hidden"
	if !strings.Contains(output, expected) {
		t.Errorf("Expected output to contain %q, got %q", expected, output)
	}
}

func TestVB6InterfaceDispatchInsideImplementsScope(t *testing.T) {
	code := `
		Class IFoo
			Sub DoSomething()
			End Sub
		End Class

		Class Foo
			Implements IFoo
			Sub IFoo_DoSomething()
				Response.Write "Foo.DoSomething called"
			End Sub
		End Class

		Class Bar
			Implements IFoo
			Private m_Foo As IFoo

			Sub Init()
				Set m_Foo = New Foo
			End Sub

			Sub IFoo_DoSomething()
				Response.Write "Bar.DoSomething calling m_Foo..."
				m_Foo.DoSomething
				Response.Write " | Bar.DoSomething done"
			End Sub
		End Class

		Dim barObj
		Set barObj = New Bar
		barObj.Init

		Dim bar As IFoo
		Set bar = barObj
		bar.DoSomething
	`

	output, err := runVBScriptTest(code)
	if err != nil {
		t.Fatalf("Execution failed: %v", err)
	}

	expected := "Bar.DoSomething calling m_Foo...Foo.DoSomething called | Bar.DoSomething done"
	if !strings.Contains(output, expected) {
		t.Errorf("Expected output to contain %q, got %q", expected, output)
	}
}

func TestVB6ConcreteClassDispatch(t *testing.T) {
	code := `
		Class Dog
			Function Speak()
				Speak = "Woof!"
			End Function
		End Class

		Dim d As Dog
		Set d = New Dog
		Response.Write d.Speak()
	`

	output, err := runVBScriptTest(code)
	if err != nil {
		t.Fatalf("Execution failed: %v", err)
	}

	expected := "Woof!"
	if !strings.Contains(output, expected) {
		t.Errorf("Expected output to contain %q, got %q", expected, output)
	}
}

func TestVB6UninitializedTypedIsNothing(t *testing.T) {
	code := `
		Class Dog
		End Class

		Class Cat
			Public MyDog As Dog
		End Class

		Dim d As Dog
		If d Is Nothing Then
			Response.Write "GlobalIsNothing "
		End If

		Sub TestLocal()
			Dim localD As Dog
			If localD Is Nothing Then
				Response.Write "LocalIsNothing "
			End If
		End Sub
		Call TestLocal()

		Dim c
		Set c = New Cat
		If c.MyDog Is Nothing Then
			Response.Write "MemberIsNothing"
		Else
			Response.Write "MemberNotNothing"
		End If
	`

	output, err := runVBScriptTest(code)
	if err != nil {
		t.Fatalf("Execution failed: %v", err)
	}

	expected := "GlobalIsNothing LocalIsNothing MemberIsNothing"
	if !strings.Contains(output, expected) {
		t.Errorf("Expected output to contain %q, got %q", expected, output)
	}
}

func TestVB6InterfaceTypedReturn(t *testing.T) {
	code := `
		Class IAnimal
			Function Clone()
			End Function
			Function Clone2()
			End Function
		End Class

		Class Dog
			Implements IAnimal
			
			Public Function IAnimal_Clone() As IAnimal
				Set IAnimal_Clone = Me
			End Function

			Public Function IAnimal_Clone2 As IAnimal
				Set IAnimal_Clone2 = Me
			End Function

			Public Function Speak()
				Speak = "Woof!"
			End Function
		End Class

		Dim a As IAnimal
		Set a = New Dog
		
		Dim b As IAnimal
		Set b = a.Clone()

		Dim c As IAnimal
		Set c = a.Clone2()
		
		Dim d
		Set d = b
		Response.Write d.Speak()

		Dim e
		Set e = c
		Response.Write " " & e.Speak()
	`

	output, err := runVBScriptTest(code)
	if err != nil {
		t.Fatalf("Execution failed: %v", err)
	}

	expected := "Woof! Woof!"
	if !strings.Contains(output, expected) {
		t.Errorf("Expected output to contain %q, got %q", expected, output)
	}
}
