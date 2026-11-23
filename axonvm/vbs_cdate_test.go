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

// CDate follows VBScript: numbers are OLE dates (days since 12/30/1899), Empty is
// the zero date, and text that is not a date raises Type mismatch. It used to
// return a garbage date (8/30/1754) for all of these without an error.
func TestVBSCDateConversions(t *testing.T) {
	out := runVBSAndGetOutput(t, `<%
On Error Resume Next
For Each s In Array(45000, 45000.5, -1.25, Empty, "2026/10/01", "12:30", "abc", "", "2026/13/45")
Err.Clear
d = "-"
d = CDate(s)
Response.Write d & "/" & (Err.Number <> 0) & "|"
Next
%>`)
	want := "3/15/2023/False|3/15/2023 12:00:00 PM/False|12/29/1899 6:00:00 AM/False|12:00:00 AM/False|10/1/2026/False|12:30:00 PM/False|-/True|-/True|-/True|"
	if out != want {
		t.Fatalf("got  %q\nwant %q", out, want)
	}
}
