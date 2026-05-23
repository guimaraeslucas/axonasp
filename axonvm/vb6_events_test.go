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
