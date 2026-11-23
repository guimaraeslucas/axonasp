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

// TestVBSSplitDefaultsAndOptions checks Split's default space delimiter,
// empty input, empty delimiter, limit and vbTextCompare.
func TestVBSSplitDefaultsAndOptions(t *testing.T) {
	src := `<%
a = Split("Mozilla/5.0 Chrome/131.0 Safari/537")
Response.Write UBound(a) & ":" & a(2) & "|"
Response.Write UBound(Split("")) & "|" & UBound(Split("", ",")) & "|"
b = Split("abc", "")
Response.Write UBound(b) & ":" & b(0) & "|"
c = Split("a,b,c", ",", 2)
Response.Write UBound(c) & ":" & c(1) & "|"
d = Split("1x2X3", "x", -1, 1)
Response.Write UBound(d) & ":" & d(2) & "|"
e = Split("1x2X3", "x")
Response.Write UBound(e) & ":" & e(1)
%>`
	got := runVBSAndGetOutput(t, src)
	want := "2:Safari/537|-1|-1|0:abc|1:b,c|2:3|1:2X3"
	if got != want {
		t.Fatalf("got %q, want %q", got, want)
	}
}
