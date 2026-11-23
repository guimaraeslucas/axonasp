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

// TestVBSNumericStringComparesWithNumber checks that a numeric string compared
// with a number is compared numerically, not as text ("131" > 82).
func TestVBSNumericStringComparesWithNumber(t *testing.T) {
	src := `<%
v = "131"
Response.Write "" & (v > 82) & "|" & (v < 82) & "|" & (82 < v) & "|" & (v >= 131) & "|" & (v = 131) & "|"
Response.Write "" & ("9" < 10) & "|" & (" 2.5 " > 2) & "|" & ("abc" > 1) & "|"
If v > 82 Then Response.Write "yes" Else Response.Write "no"
%>`
	got := runVBSAndGetOutput(t, src)
	want := "True|False|True|True|True|True|True|True|yes"
	if got != want {
		t.Fatalf("got %q, want %q", got, want)
	}
}
