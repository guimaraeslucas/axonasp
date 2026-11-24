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

// TestADODBAbsolutePageMovesCursor checks that setting AbsolutePage moves the
// cursor to the first record of that page, as ADO does for paging loops.
func TestADODBAbsolutePageMovesCursor(t *testing.T) {
	src := `<%
Set rs = Server.CreateObject("ADODB.Recordset")
rs.Fields.Append "id", 3
rs.Open
For i = 1 To 25
  rs.AddNew
  rs("id") = i
  rs.Update
Next
rs.PageSize = 10
Response.Write rs.PageCount & "|"
rs.MoveFirst
rs.AbsolutePage = "3"
Response.Write rs("id") & "|" & rs.AbsolutePage & "|" & rs.AbsolutePosition & "|"
n = 0
While n < rs.PageSize And Not rs.EOF
  n = n + 1
  rs.MoveNext
Wend
Response.Write n & "|"
rs.AbsolutePage = 2
Response.Write rs("id") & "|" & rs.AbsolutePage
%>`
	got := runVBSAndGetOutput(t, src)
	want := "3|21|3|21|5|11|2"
	if got != want {
		t.Fatalf("got %q, want %q", got, want)
	}
}
