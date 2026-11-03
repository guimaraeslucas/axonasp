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

func TestVBSDatePartsOfNullAreNull(t *testing.T) {
	src := `<%
n = Null
Response.Write IsNull(Year(n)) & IsNull(Month(n)) & IsNull(Day(n)) & IsNull(Hour(n)) & "|"
Response.Write IsNull(Weekday(n)) & IsNull(DatePart("m", n)) & IsNull(DateAdd("n", 180, n)) & "|"
Response.Write IsNull(Left(n, 2)) & IsNull(Right(n, 2)) & IsNull(Mid(n, 2)) & "|"
' EUformatDDMMMYY in crazy-crs: an empty date renders as "--"
s = Day(n) & "-" & MonthName(Month(n)) & "-" & Right(Year(n), 2)
Response.Write s & "|"
Set rs = CreateObject("ADODB.Recordset")
rs.Fields.Append "d", 7, , 32
rs.Open
rs.AddNew
rs("d") = Null
rs.Update
rs.MoveFirst
Response.Write IsNull(Month(rs("d"))) & "|" & Month(CDate("2026-11-03"))
%>`
	want := "TrueTrueTrueTrue|TrueTrueTrue|TrueTrueTrue|--|True|11"
	if got := runVBSAndGetOutput(t, src); got != want {
		t.Fatalf("got %q, want %q", got, want)
	}
}
