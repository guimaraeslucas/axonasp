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

func TestJScriptArrayCalledWithoutNew(t *testing.T) {
	cases := []struct{ src, want string }{
		{`<script language="javascript" runat="server">function f(){ var a = Array(3); return typeof a + "/" + a.length; }</script><%=f()%>`, "object/3"},
		{`<script language="javascript" runat="server">function f(){ var a = Array(1, 2); return a.length + "/" + a[1]; }</script><%=f()%>`, "2/2"},
		{`<script language="javascript" runat="server">function f(){ var a = Array(); a[2] = 7; return a.length; }</script><%=f()%>`, "3"},
	}
	for _, tc := range cases {
		if got := runVBSAndGetOutput(t, tc.src); got != tc.want {
			t.Errorf("%s => %q, want %q", tc.src, got, tc.want)
		}
	}
}
