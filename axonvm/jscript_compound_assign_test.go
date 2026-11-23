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

func TestJScriptCompoundBitwiseAssign(t *testing.T) {
	cases := []struct{ src, want string }{
		{`<script language="javascript" runat="server">function f(){ var x = 12; x |= 3; return x; }</script><%=f()%>`, "15"},
		{`<script language="javascript" runat="server">function f(){ var x = 12; x &= 6; return x; }</script><%=f()%>`, "4"},
		{`<script language="javascript" runat="server">function f(){ var x = 12; x ^= 5; return x; }</script><%=f()%>`, "9"},
		{`<script language="javascript" runat="server">function f(){ var x = 3; x <<= 4; return x; }</script><%=f()%>`, "48"},
		{`<script language="javascript" runat="server">function f(){ var x = -16; x >>= 2; return x; }</script><%=f()%>`, "-4"},
		{`<script language="javascript" runat="server">function f(){ var x = -16; x >>>= 28; return x; }</script><%=f()%>`, "15"},
		{`<script language="javascript" runat="server">function f(){ var x; x |= 6; return x; }</script><%=f()%>`, "6"},
		{`<script language="javascript" runat="server">function f(){ var a = [1,2]; a[0] += 3; return a[0]; }</script><%=f()%>`, "4"},
		{`<script language="javascript" runat="server">function f(){ var a = [12]; var r = (a[0] ^= 5); return r + "/" + a[0]; }</script><%=f()%>`, "9/9"},
		{`<script language="javascript" runat="server">function f(){ var a = [1,2,4]; var i = 0; a[i++] |= 8; return a + "/" + i; }</script><%=f()%>`, "9,2,4/1"},
		{`<script language="javascript" runat="server">function f(){ var a = []; a[3] |= 5; return a.length + "/" + a[3]; }</script><%=f()%>`, "4/5"},
		{`<script language="javascript" runat="server">function f(){ var o = {y: 1}; o.y |= 6; return o.y; }</script><%=f()%>`, "7"},
		{`<script language="javascript" runat="server">function f(){ var o = {k: 3}; var r = (o["k"] <<= 2); return r + "/" + o.k; }</script><%=f()%>`, "12/12"},
	}
	for _, tc := range cases {
		if got := runVBSAndGetOutput(t, tc.src); got != tc.want {
			t.Errorf("%s => %q, want %q", tc.src, got, tc.want)
		}
	}
}
