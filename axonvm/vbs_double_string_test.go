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
 * made available under the same license terms.
 */

import "testing"

// TestVBScriptDoubleToString checks that Doubles convert to strings the way VBScript
// does (15 significant digits, E notation only from 1E+15), both at runtime and in
// compile-time constant folding. Go's %g printed CDbl(45620382) as "4.5620382e+07".
func TestVBScriptDoubleToString(t *testing.T) {
	src := `<% Dim d : d = CDbl(45620382)
Response.Write d & "|" & CStr(CDbl(1234567890123)) & "|" & CDbl(100000000) & "|" & (0.1 + 0.2) & "|" & 2 ^ 53 & "|" & 3.5 & "|" & ("x" & 45620382.0) %>`
	want := "45620382|1234567890123|100000000|0.3|9.00719925474099E+15|3.5|x45620382"
	if got := runVBSAndGetOutput(t, src); got != want {
		t.Fatalf("got %q, want %q", got, want)
	}
}
