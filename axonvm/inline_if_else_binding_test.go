package axonvm

/*
 * AxonASP Server
 * Copyright (C) 2026 G3pix Ltda. All rights reserved.
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 */

import (
	"strings"
	"testing"

	"g3pix.com.br/axonasp/v2/vbscript"
)

// A single-line If is complete at an ASP tag boundary, so an Else or ElseIf in the next
// script block never attaches to it: with an enclosing block If open it binds there, and
// with none it is a syntax error. Whether the boundary is empty makes no difference.
//
// Every expectation in this file is measured on IIS 10 (saborrie/iis-oracle):
//
//	binding to the outer block If:   https://github.com/saborrie/iis-oracle/actions/runs/42511937887
//	detached Else/ElseIf rejected:   https://github.com/saborrie/iis-oracle/actions/runs/42512053123
//	empty boundary changes nothing:  https://github.com/saborrie/iis-oracle/actions/runs/42512101127
//	End If after a same-line chain:  https://github.com/saborrie/iis-oracle/actions/runs/42512373768
//
// Classic ASP renders one of two markup fragments this way routinely, with a detail
// tweak as a single-line If at the end of the first fragment.

// TestTagBoundaryElseBindsToTheOuterBlockIf runs the ifelse-outer-true/false probes:
// outer condition true prints the first fragment (inner single-line If stays false),
// outer condition false prints the Else fragment.
func TestTagBoundaryElseBindsToTheOuterBlockIf(t *testing.T) {
	page := `<%
Dim a, y
a = %TRUE%
y = 0
If a = 1 Then %>x<% If a = 2 Then y = 1 %><% Else %>z<% End If
%>[y=<%=y%>]`

	for _, tc := range []struct {
		name, a, want string
	}{
		{"outer true takes the first fragment", "1", "x[y=0]"},
		{"outer false takes the else fragment", "3", "z[y=0]"},
	} {
		t.Run(tc.name, func(t *testing.T) {
			source := strings.ReplaceAll(page, "%TRUE%", tc.a)
			if got := runVBSAndGetOutput(t, source); got != tc.want {
				t.Fatalf("got %q, want %q", got, tc.want)
			}
		})
	}
}

// TestTagBoundaryElseIfBindsToTheOuterBlockIf is the ElseIf spelling, both directions,
// mirroring the ifelseif-tag-boundary probe: the first direction runs the outer If arm
// and its inner single-line If (y becomes 2), the second runs the outer ElseIf arm.
func TestTagBoundaryElseIfBindsToTheOuterBlockIf(t *testing.T) {
	source := `<%
Dim a, b, y
a = 2 : b = 0 : y = 0
If a <> 0 Then %>A<% If a <> 1 Then y = 2 %><% ElseIf b <> 0 Then %>B<% End If
%>[first y=<%=y%>]<%
a = 0 : b = 5 : y = 0
If a <> 0 Then %>A<% If a <> 1 Then y = 2 %><% ElseIf b <> 0 Then %>B<% End If
%>[second y=<%=y%>]`

	want := "A[first y=2]B[second y=0]"
	if got := runVBSAndGetOutput(t, source); got != want {
		t.Fatalf("got %q, want %q", got, want)
	}
}

// TestTagBoundaryElseBindsToTheOuterBlockIfInsideALoop keeps the binding stable inside
// a While body, where the old absorption surfaced its error at the Wend.
func TestTagBoundaryElseBindsToTheOuterBlockIfInsideALoop(t *testing.T) {
	source := `<%
Dim n, hours, days, y
hours = 2 : days = 0 : y = 0
n = 0
While n < 2
	If hours <> 0 Then %><%=hours%>h<% If hours <> 1 Then y = y + 1 %><% ElseIf days <> 0 Then %><%=days%>d<% End If
	n = n + 1
Wend
%>[y=<%=y%>]`

	want := "2h2h[y=2]"
	if got := runVBSAndGetOutput(t, source); got != want {
		t.Fatalf("got %q, want %q", got, want)
	}
}

// TestDetachedElseAfterACompletedSingleLineIfIsRejected pins the error cases: with no
// enclosing block If, the Else/ElseIf in the next script block is a syntax error on
// IIS (800A03EA at the branch keyword), HTML between the tags or not.
func TestDetachedElseAfterACompletedSingleLineIfIsRejected(t *testing.T) {
	for _, tc := range []struct {
		name, source string
	}{
		{
			"else with html between the tags",
			`<% Dim a, y : a = 1 : y = 0 : If a = 1 Then y = 1 %>x<% Else %>z<% End If %>`,
		},
		{
			"else across an empty boundary",
			`<% Dim a, y : a = 1 : y = 0 : If a = 1 Then y = 1 %><% Else %>z<% End If %>`,
		},
		{
			"elseif with html between the tags",
			`<% Dim a, y : a = 1 : y = 0 : If a = 1 Then y = 1 %>x<% ElseIf a = 2 Then %>z<% End If %>`,
		},
	} {
		t.Run(tc.name, func(t *testing.T) {
			requireCompileError(t, tc.source, vbscript.SyntaxError)
		})
	}
}

// TestTagBoundaryEndIfAfterACompleteSameLineChainIsRejected: a same-line If/Else chain
// is complete at the tag boundary exactly like the plain single-line If, so the End If
// in the next script block is an orphan. IIS rejects it with 800A0400 "Expected
// statement" at the End If; the engine reports its orphan End If error.
func TestTagBoundaryEndIfAfterACompleteSameLineChainIsRejected(t *testing.T) {
	requireOrphanEndIfError(t,
		`<% Dim a, y : a = 3 : y = 0 : If a = 1 Then y = 1 Else y = 2 %><% End If %>`)
}
