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

import "testing"

// TestSingleLineIfColonReproduction validates the reproduction case where a colon
// follows 'Then' in a single-line If statement without requiring End If.
func TestSingleLineIfColonReproduction(t *testing.T) {
	source := `<%
Dim a, b
a = "x"
if a <> "" then : b = "set" : a = "done"
Response.Write "b=" & b & " a=" & a
%>`
	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	out := runVBSAndGetOutput(t, source)
	expected := "b=set a=done"
	if out != expected {
		t.Fatalf("unexpected output: got %q, want %q", out, expected)
	}
}

// TestSingleLineIfColonMultipleStatements verifies multiple colon-separated statements on single-line If.
func TestSingleLineIfColonMultipleStatements(t *testing.T) {
	source := `<%
If True Then : Response.Write "A" : Response.Write "B"
%>`
	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	out := runVBSAndGetOutput(t, source)
	expected := "AB"
	if out != expected {
		t.Fatalf("unexpected output: got %q, want %q", out, expected)
	}
}

// TestSingleLineIfColonVariableWhitespace verifies single-line If with variable whitespace before/after colon.
func TestSingleLineIfColonVariableWhitespace(t *testing.T) {
	source := `<%
Dim a
a = 0
If True Then   :   a = 1
Response.Write "a=" & a
%>`
	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	out := runVBSAndGetOutput(t, source)
	expected := "a=1"
	if out != expected {
		t.Fatalf("unexpected output: got %q, want %q", out, expected)
	}
}

// TestBlockIfWithColonsInStatements verifies that block If statements containing colons in inner statements
// continue to parse and execute properly as blocks without regression.
func TestBlockIfWithColonsInStatements(t *testing.T) {
	source := `<%
Dim a, b, c, d
a = 0 : b = 0 : c = 0 : d = 0
If True Then
    a = 10 : b = 20
    If True Then : c = 30 : d = 40
    Response.Write "a=" & a & " b=" & b & " c=" & c & " d=" & d
End If
%>`
	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	out := runVBSAndGetOutput(t, source)
	expected := "a=10 b=20 c=30 d=40"
	if out != expected {
		t.Fatalf("unexpected output: got %q, want %q", out, expected)
	}
}
