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
	"errors"
	"testing"

	"g3pix.com.br/axonasp/vbscript"
)

// TestCompilerUnexpectedTokenUsesVBScriptMetadata verifies compiler failures expose VBScript-compatible metadata.
func TestCompilerUnexpectedTokenUsesVBScriptMetadata(t *testing.T) {
	compiler := NewASPCompiler(`<%= * %>`)
	compiler.SetSourceName("/tests/compiler_error.asp")

	err := compiler.Compile()
	if err == nil {
		t.Fatalf("expected compile error")
	}

	var syntaxErr *vbscript.VBSyntaxError
	if !errors.As(err, &syntaxErr) {
		t.Fatalf("expected VBScript syntax error, got %T", err)
	}

	if syntaxErr.Code != vbscript.SyntaxError {
		t.Fatalf("unexpected code: got %d want %d", syntaxErr.Code, vbscript.SyntaxError)
	}
	if syntaxErr.File != "/tests/compiler_error.asp" {
		t.Fatalf("unexpected file: %q", syntaxErr.File)
	}
	if syntaxErr.Category != "VBScript compilation" {
		t.Fatalf("unexpected category: %q", syntaxErr.Category)
	}
	if syntaxErr.Source != "VBScript compilation error" {
		t.Fatalf("unexpected source: %q", syntaxErr.Source)
	}
	if syntaxErr.Description != vbscript.SyntaxError.String() {
		t.Fatalf("unexpected description: %q", syntaxErr.Description)
	}
	if syntaxErr.ASPDescription == "" {
		t.Fatalf("expected ASP description detail")
	}
	if syntaxErr.Number != vbscript.HRESULTFromVBScriptCode(vbscript.SyntaxError) {
		t.Fatalf("unexpected number: got %d want %d", syntaxErr.Number, vbscript.HRESULTFromVBScriptCode(vbscript.SyntaxError))
	}
	if syntaxErr.Column != 7 {
		t.Fatalf("unexpected column: got %d want 7", syntaxErr.Column)
	}

	aspErr := CompilerErrorToASPError(err, "/tests/compiler_error.asp")
	if aspErr.ASPCode != int(vbscript.SyntaxError) {
		t.Fatalf("unexpected asp code: got %d want %d", aspErr.ASPCode, vbscript.SyntaxError)
	}
	if aspErr.File != "/tests/compiler_error.asp" {
		t.Fatalf("unexpected asp file: %q", aspErr.File)
	}
}

// TestCompilerConstAssignmentFails verifies assigning to a Const raises VBScript IllegalAssignment.
func TestCompilerConstAssignmentFails(t *testing.T) {
	compiler := NewASPCompiler(`<%
Const MyNum = 100
MyNum = 200
%>`)

	err := compiler.Compile()
	if err == nil {
		t.Fatalf("expected compile error")
	}

	var syntaxErr *vbscript.VBSyntaxError
	if !errors.As(err, &syntaxErr) {
		t.Fatalf("expected VBScript syntax error, got %T", err)
	}

	if syntaxErr.Code != vbscript.IllegalAssignment {
		t.Fatalf("unexpected code: got %d want %d", syntaxErr.Code, vbscript.IllegalAssignment)
	}
}

func TestCompilerRejectsOnErrorGotoLabel(t *testing.T) {
	compiler := NewASPCompiler(`<%
On Error GoTo MyHandler
%>`)

	err := compiler.Compile()
	if err == nil {
		t.Fatalf("expected compile error")
	}

	var syntaxErr *vbscript.VBSyntaxError
	if !errors.As(err, &syntaxErr) {
		t.Fatalf("expected VBScript syntax error, got %T", err)
	}
	if syntaxErr.Code != vbscript.SyntaxError {
		t.Fatalf("unexpected code: got %d want %d", syntaxErr.Code, vbscript.SyntaxError)
	}
}
