package axonvm

/*
 * AxonASP Server
 * Copyright (C) 2026 G3pix Ltda. All rights reserved.
 *
 * Developed by Yuri Eisman (@yeisman)
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

import "testing"

// If a > b / If a < b compile to fused compare-jumps; they must compare like
// the plain operators (a numeric string vs a number compares numerically).
func TestVBSIfCompareStringToNumber(t *testing.T) {
	src := `<%
s = "3063" : g = FormatNumber(39815 / 13, 0) : y = 0
If s > 0 Then Response.Write "1" Else Response.Write "0"
If g > 0 Then Response.Write "1" Else Response.Write "0"
If s > y Then Response.Write "1" Else Response.Write "0"
If 0 < s Then Response.Write "1" Else Response.Write "0"
If s < 0 Then Response.Write "1" Else Response.Write "0"
If 2.5 > 2 Then Response.Write "1" Else Response.Write "0"
If 2 < 2.5 Then Response.Write "1" Else Response.Write "0"
If "b" > "a" Then Response.Write "1" Else Response.Write "0"
If s > 0 Then
	Response.Write "1"
Else
	Response.Write "0"
End If
If g > 0 Then
	Response.Write "1"
Else
	Response.Write "0"
End If
If s > y Then
	Response.Write "1"
Else
	Response.Write "0"
End If
If 0 < s Then
	Response.Write "1"
Else
	Response.Write "0"
End If
If s < 0 Then
	Response.Write "1"
Else
	Response.Write "0"
End If
If 2.5 > 2 Then
	Response.Write "1"
Else
	Response.Write "0"
End If
If 2 < 2.5 Then
	Response.Write "1"
Else
	Response.Write "0"
End If
If "b" > "a" Then
	Response.Write "1"
Else
	Response.Write "0"
End If
Response.Write "|" & (s > 0) & (s < 0)
%>`
	want := "1111011111110111|TrueFalse"
	if got := runVBSAndGetOutput(t, src); got != want {
		t.Fatalf("got %q, want %q", got, want)
	}
}
