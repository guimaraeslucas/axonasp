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
	"strings"
	"testing"
)

func runVBScriptTest(source string) (string, error) {
	if !strings.Contains(source, "<%") {
		source = "<% " + source + " %>"
	}
	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		return "", err
	}
	vm := NewVMFromCompiler(compiler)
	var buf bytes.Buffer
	// We need a mock host that captures output
	host := NewMockHost()
	host.SetOutput(&buf)
	vm.SetHost(host)
	if err := vm.Run(); err != nil {
		return "", err
	}
	host.Response().Flush()
	return buf.String(), nil
}

func TestVB6Events(t *testing.T) {
	code := `
		Class MySource
			Event OnClick(val)
			Sub DoClick(v)
				RaiseEvent OnClick(v)
			End Sub
		End Class

		Class MySink
			Public Result
			Private WithEvents m_src

			Sub Class_Initialize()
				Set m_src = New MySource
			End Sub

			Sub m_src_OnClick(v)
				Result = "Clicked: " & v
			End Sub

			Sub ClickIt(v)
				m_src.DoClick v
			End Sub
		End Class

		Set sink = New MySink
		sink.ClickIt "Hello"
		Response.Write sink.Result
	`

	output, err := runVBScriptTest(code)
	if err != nil {
		t.Fatalf("Execution failed: %v", err)
	}

	expected := "Clicked: Hello"
	if !strings.Contains(output, expected) {
		t.Errorf("Expected output to contain %q, got %q", expected, output)
	}
}

func TestVB6GlobalWithEvents(t *testing.T) {
	code := `
		Class MySource
			Event OnChange()
			Sub Trigger()
				RaiseEvent OnChange()
			End Sub
		End Class

		Dim WithEvents g_src
		Dim g_result
		g_result = "Initial"

		Sub g_src_OnChange()
			g_result = "Changed"
		End Sub

		Set g_src = New MySource
		g_src.Trigger
		Response.Write g_result
	`

	output, err := runVBScriptTest(code)
	if err != nil {
		t.Fatalf("Execution failed: %v", err)
	}

	expected := "Changed"
	if !strings.Contains(output, expected) {
		t.Errorf("Expected output to contain %q, got %q", expected, output)
	}
}

func TestVB6EventsReassignment(t *testing.T) {
	code := `
		Class MySource
			Public ID
			Event OnNotify(msg)
			Sub Notify(m)
				RaiseEvent OnNotify("Source " & ID & ": " & m)
			End Sub
		End Class

		Dim WithEvents src
		Dim result
		
		Sub src_OnNotify(m)
			result = m
		End Sub

		Set src = New MySource
		src.ID = 1
		
		Set src2 = New MySource
		src2.ID = 2
		
		Set src = src2 ' Should unbind from src1, bind to src2
		
		src2.Notify "Hello"
		Response.Write result
	`

	output, err := runVBScriptTest(code)
	if err != nil {
		t.Fatalf("Execution failed: %v", err)
	}

	expected := "Source 2: Hello"
	if !strings.Contains(output, expected) {
		t.Errorf("Expected output to contain %q, got %q", expected, output)
	}
}

// TestVB6EventTypedParam verifies that an Event declaration with As Type on a parameter
// compiles and executes correctly (issue: parser previously rejected 'As' in event params).
func TestVB6EventTypedParam(t *testing.T) {
	code := `
		Class EventSource
			Event Bar(x As String)

			Public Function DoWork
				RaiseEvent Bar("hello")
			End Function
		End Class

		Class EventHandler
			Dim WithEvents src As EventSource

			Public Function Run
				Set src = New EventSource
				src.DoWork
			End Function

			Sub src_Bar(x As String)
				Response.Write "Received: " & x
			End Sub
		End Class

		Dim h As EventHandler
		Set h = New EventHandler
		h.Run
	`

	output, err := runVBScriptTest(code)
	if err != nil {
		t.Fatalf("Execution failed: %v", err)
	}

	expected := "Received: hello"
	if !strings.Contains(output, expected) {
		t.Errorf("Expected output to contain %q, got %q", expected, output)
	}
}

// TestVB6EventUntypedParam verifies backward compatibility: an Event declaration
// without As Type continues to compile and execute correctly.
func TestVB6EventUntypedParam(t *testing.T) {
	code := `
		Class EventSource
			Event OnComplete(status)
			Sub Fire
				RaiseEvent OnComplete("done")
			End Sub
		End Class

		Dim WithEvents src
		Dim result
		result = ""

		Sub src_OnComplete(s)
			result = s
		End Sub

		Set src = New EventSource
		src.Fire
		Response.Write result
	`

	output, err := runVBScriptTest(code)
	if err != nil {
		t.Fatalf("Execution failed: %v", err)
	}

	expected := "done"
	if !strings.Contains(output, expected) {
		t.Errorf("Expected output to contain %q, got %q", expected, output)
	}
}

// TestVB6EventMixedParams verifies that an event with a mix of typed and untyped
// parameters, as well as ByVal/ByRef modifiers, compiles without error.
func TestVB6EventMixedParams(t *testing.T) {
	code := `
		Class EventSource
			Event MixedEvent(ByVal str As String, obj, ByRef num As Integer)
			Sub Fire
				RaiseEvent MixedEvent("test", Nothing, 42)
			End Sub
		End Class

		Dim WithEvents src
		Dim result
		result = ""

		Sub src_MixedEvent(ByVal str As String, obj, ByRef num As Integer)
			result = str
		End Sub

		Set src = New EventSource
		src.Fire
		Response.Write result
	`

	output, err := runVBScriptTest(code)
	if err != nil {
		t.Fatalf("Compilation failed for typed event params: %v", err)
	}

	expected := "test"
	if !strings.Contains(output, expected) {
		t.Errorf("Expected output to contain %q, got %q", expected, output)
	}
}

// TestVB6EventNoParams verifies that an event with no parameters still compiles.
func TestVB6EventNoParams(t *testing.T) {
	code := `
		Class MySource
			Event OnChange()
			Sub Trigger
				RaiseEvent OnChange
			End Sub
		End Class

		Class MySink
			Private WithEvents src
			Public Result

			Sub Class_Initialize
				Set src = New MySource
			End Sub

			Sub src_OnChange()
				Result = "changed"
			End Sub
		End Class

		Dim s As MySink
		Set s = New MySink
		Response.Write "ok"
	`

	output, err := runVBScriptTest(code)
	if err != nil {
		t.Fatalf("Compilation failed for no-param event: %v", err)
	}

	if !strings.Contains(output, "ok") {
		t.Errorf("Expected output to contain 'ok', got %q", output)
	}
}
