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

// VBScript parses dates with English month names in the formats classic ASP
// pages print ("05-Aug-26", "5 Aug 2026", "Aug 5, 2026"), with an optional time,
// and maps two-digit years 00-29 to 20xx and 30-99 to 19xx.
func TestVBSMonthNameDateStrings(t *testing.T) {
	out := runVBSAndGetOutput(t, `<%
For Each s In Array("05-Aug-26", "05-aug-2026", "5 Aug 2026", "Aug 5, 2026", "August 5, 2026", "05-Aug-26 03:30", "05-Aug-45", "05-Aug-29")
Response.Write IsDate(s) & " " & Year(s) & "-" & Month(s) & "-" & Day(s) & " " & Hour(s) & ":" & Minute(s) & "|"
Next
%>`)
	want := "True 2026-8-5 0:0|True 2026-8-5 0:0|True 2026-8-5 0:0|True 2026-8-5 0:0|True 2026-8-5 0:0|True 2026-8-5 3:30|True 1945-8-5 0:0|True 2029-8-5 0:0|"
	if out != want {
		t.Fatalf("got  %q\nwant %q", out, want)
	}
}
