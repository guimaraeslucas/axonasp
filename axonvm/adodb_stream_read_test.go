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

// b = stream.Read without parentheses is a method call that reads to the end
// (it returned Empty); the text-to-bytes helper below is a common classic ASP idiom.
func TestADODBStreamReadWithoutParens(t *testing.T) {
	src := `<%
Set st = Server.CreateObject("ADODB.Stream")
st.Type = 2 : st.CharSet = "us-ascii" : st.Open : st.WriteText "123_1"
st.Position = 0 : st.Type = 1 : st.Position = 0
b = st.Read
Response.Write LenB(b) & "|" & AscB(MidB(b, 1, 1)) & "|" & st.EOS
%>`
	if got := runVBSAndGetOutput(t, src); got != "5|49|True" {
		t.Fatalf("got %q, want %q", got, "5|49|True")
	}
}
