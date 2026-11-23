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

// In VBScript True is -1 when used as a number (IIS: CInt(True) = -1, True = 1 is False).
// Classic ASP pages test bit columns with `If rs("flag") = 1`, which is False on IIS.
func TestVBSBooleanNumericValueIsMinusOne(t *testing.T) {
	out := runVBSAndGetOutput(t, `<%
x = True : y = False
Response.Write "" & (x = 1) & "|" & (x = -1) & "|" & CInt(x) & "|" & CLng(x) & "|" & CDbl(x) & "|" & (x + 0) & "|" & (x * 2) & "|" & (x - 0)
Response.Write "/" & (y = 0) & "|" & Abs(x) & "|" & (1 = True) & "|" & (-1 = True) & "|" & (x And 3) & "|" & (x + x) & "|" & (x < 0) & "|" & (x > 0)
Response.Write "/" & CStr(x) & "|" & (Not 0) & "|" & (Not x) & "|" & (x Or 0) & "|" & FormatNumber(x, 0)
If x Then Response.Write "/if"
%>`)
	want := "False|True|-1|-1|-1|-1|-2|-1/True|1|False|True|3|-2|True|False/True|-1|False|-1|-1/if"
	if out != want {
		t.Fatalf("got  %q\nwant %q", out, want)
	}
}
