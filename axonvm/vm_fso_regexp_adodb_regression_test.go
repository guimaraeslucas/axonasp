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
	"runtime"
	"strconv"
	"strings"
	"testing"

	"g3pix.com.br/axonasp/vbscript"
)

// TestASPForEachFSODrivesCollection verifies For Each over fso.Drives returns drive objects.
func TestASPForEachFSODrivesCollection(t *testing.T) {
	source := `<%
Dim fso, d, enumCount, firstLetter
Set fso = Server.CreateObject("Scripting.FileSystemObject")
enumCount = 0
firstLetter = ""
For Each d In fso.Drives
    enumCount = enumCount + 1
    If firstLetter = "" Then
        firstLetter = d.DriveLetter
    End If
Next
Response.Write fso.Drives.Count & "|" & enumCount & "|" & firstLetter
%>`

	actual := runASPAndCollectOutput(t, source)
	parts := strings.Split(actual, "|")
	if len(parts) != 3 {
		t.Fatalf("unexpected output format: %q", actual)
	}
	declaredCount, err := strconv.Atoi(parts[0])
	if err != nil {
		t.Fatalf("invalid declared count %q: %v", parts[0], err)
	}
	enumCount, err := strconv.Atoi(parts[1])
	if err != nil {
		t.Fatalf("invalid enumeration count %q: %v", parts[1], err)
	}
	firstLetter := parts[2]

	if declaredCount != enumCount {
		t.Fatalf("drives count mismatch: declared=%d enumerated=%d output=%q", declaredCount, enumCount, actual)
	}
	if runtime.GOOS == "windows" {
		if enumCount < 1 {
			t.Fatalf("expected at least one drive on Windows, got output %q", actual)
		}
		if len(firstLetter) != 1 {
			t.Fatalf("expected first drive letter on Windows, got %q", firstLetter)
		}
		return
	}
	if actual != "1|1|C" {
		t.Fatalf("unexpected non-Windows fallback drive output: got %q want %q", actual, "1|1|C")
	}
}

// TestASPFSOTextStreamSkipLine verifies SkipLine advances to the next line before ReadLine.
func TestASPFSOTextStreamSkipLine(t *testing.T) {
	source := `<%
Dim fso, p, ts
Set fso = Server.CreateObject("Scripting.FileSystemObject")
p = Server.MapPath("/skipline.txt")
Set ts = fso.CreateTextFile(p, True)
ts.WriteLine "A"
ts.WriteLine "B"
ts.WriteLine "C"
ts.Close

Set ts = fso.OpenTextFile(p, 1)
ts.SkipLine
Response.Write ts.Line & "|" & ts.ReadLine
ts.Close
%>`

	actual := runASPAndCollectOutput(t, source)
	if actual != "2|B" {
		t.Fatalf("unexpected SkipLine behavior: got %q want %q", actual, "2|B")
	}
}

// TestASPForEachRegExpSubMatches verifies For Each over Match.SubMatches yields captured strings.
func TestASPForEachRegExpSubMatches(t *testing.T) {
	source := `<%
Dim rx, matches, one, sm, out
Set rx = New RegExp
rx.Pattern = "([a-z]+)-([0-9]+)"
rx.Global = False
Set matches = rx.Execute("abc-123")
Set one = matches.Item(0)
out = ""
For Each sm In one.SubMatches
    out = out & CStr(sm) & ";"
Next
Response.Write out
%>`

	actual := runASPAndCollectOutput(t, source)
	if actual != "abc;123;" {
		t.Fatalf("unexpected SubMatches For Each output: got %q want %q", actual, "abc;123;")
	}
}

// TestASPADODBFieldGetChunkAppendChunkState verifies sequential GetChunk reads and AppendChunk updates.
func TestASPADODBFieldGetChunkAppendChunkState(t *testing.T) {
	source := `<%
Dim rs, fld
Set rs = CreateObject("ADODB.Recordset")
rs.Fields.Append "Chunk", 200, 255
rs.Open
rs.AddNew
Set fld = rs.Fields.Item("Chunk")
fld.AppendChunk "abc"
fld.AppendChunk "def"
Response.Write fld.GetChunk(2) & "|" & fld.GetChunk(2) & "|" & fld.GetChunk(10) & "|" & fld.Value & "|" & rs.EditMode
%>`

	actual := runASPAndCollectOutput(t, source)
	if actual != "ab|cd|ef|abcdef|1" {
		t.Fatalf("unexpected chunk state output: got %q want %q", actual, "ab|cd|ef|abcdef|1")
	}
}

// TestASPADODBFieldGetChunkPerRowCursor verifies GetChunk keeps independent cursors per field/row.
func TestASPADODBFieldGetChunkPerRowCursor(t *testing.T) {
	source := `<%
Dim rs, fld
Set rs = CreateObject("ADODB.Recordset")
rs.Fields.Append "Chunk", 200, 255
rs.Open
rs.AddNew
Set fld = rs.Fields.Item("Chunk")
fld.Value = "abcdef"
Response.Write fld.GetChunk(2)
rs.AddNew
Set fld = rs.Fields.Item("Chunk")
fld.Value = "uvwxyz"
rs.MoveFirst
Set fld = rs.Fields.Item("Chunk")
Response.Write "|" & fld.GetChunk(2)
rs.MoveNext
Set fld = rs.Fields.Item("Chunk")
Response.Write "|" & fld.GetChunk(2)
%>`

	actual := runASPAndCollectOutput(t, source)
	if actual != "ab|cd|uv" {
		t.Fatalf("unexpected per-row chunk cursor output: got %q want %q", actual, "ab|cd|uv")
	}
}

// TestASPRegExpInvalidPatternRaises5017 verifies invalid patterns raise VBScript 5017 on assignment and execution.
func TestASPRegExpInvalidPatternRaises5017(t *testing.T) {
	source := `<%
On Error Resume Next
Dim rx
Set rx = New RegExp
rx.Pattern = "(" 
Response.Write Err.Number
Err.Clear
Dim m
Set m = rx.Execute("abc")
Response.Write "|" & Err.Number
%>`

	actual := runASPAndCollectOutput(t, source)
	expected := strconv.Itoa(vbscript.HRESULTFromVBScriptCode(vbscript.RegularExpressionSyntaxError))
	if actual != expected+"|"+expected {
		t.Fatalf("unexpected RegExp invalid pattern errors: got %q want %q", actual, expected+"|"+expected)
	}
}
