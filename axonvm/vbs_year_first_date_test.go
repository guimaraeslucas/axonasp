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

// VBScript parses year-first date strings with "/" or "-", padded or not
// (Month("2026/10/01") = 10). They were not recognised as dates at all.
func TestVBSYearFirstDateStrings(t *testing.T) {
	out := runVBSAndGetOutput(t, `<%
For Each s In Array("2026/10/01", "2026/1/5", "2026-1-5", "2026-10-07", "2026/10/01 13:45")
Response.Write IsDate(s) & " " & Year(s) & "-" & Month(s) & "-" & Day(s) & " " & Hour(s) & ":" & Minute(s) & "|"
Next
%>`)
	want := "True 2026-10-1 0:0|True 2026-1-5 0:0|True 2026-1-5 0:0|True 2026-10-7 0:0|True 2026-10-1 13:45|"
	if out != want {
		t.Fatalf("got  %q\nwant %q", out, want)
	}
}
