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

func TestVBSIntArgsAcceptDoubles(t *testing.T) {
	src := `<%
w = 100
Response.Write Left("Website user", w/9) & "|"
Response.Write Right("Website user", w/25) & "|"
Response.Write Mid("Website user", 2.0, w/50) & "|"
Response.Write Left("abcdef", 2.5) & Left("abcdef", 3.5) & "|"
Response.Write Hour(DateAdd("n", CDbl(3) * 60, CDate("2026-11-03 15:04"))) & "|"
Response.Write InStr(2.0, "abcabc", "a")
%>`
	want := "Website use|user|eb|ababcd|18|4"
	if got := runVBSAndGetOutput(t, src); got != want {
		t.Fatalf("got %q, want %q", got, want)
	}
}
