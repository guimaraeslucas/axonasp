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

func TestJScriptToInt32WrapsModulo2To32(t *testing.T) {
	cases := []struct{ src, want string }{
		{`<script language="javascript" runat="server">function f(){ return 0x8f01cfea | 0; }</script><%=f()%>`, "-1895706646"},
		{`<script language="javascript" runat="server">function f(){ return 3000000000 | 0; }</script><%=f()%>`, "-1294967296"},
		{`<script language="javascript" runat="server">function f(){ return 2399260650 & 0xFF; }</script><%=f()%>`, "234"},
		{`<script language="javascript" runat="server">function f(){ return (0x8f01cfea >>> 8) & 0xFF; }</script><%=f()%>`, "207"},
		{`<script language="javascript" runat="server">function f(){ return 4294967296 | 0; }</script><%=f()%>`, "0"},
		{`<script language="javascript" runat="server">function f(){ return -1 >>> 0; }</script><%=f()%>`, "4294967295"},
		{`<script language="javascript" runat="server">function f(){ return ~0xFFFFFFFF; }</script><%=f()%>`, "0"},
	}
	for _, tc := range cases {
		if got := runVBSAndGetOutput(t, tc.src); got != tc.want {
			t.Errorf("%s => %q, want %q", tc.src, got, tc.want)
		}
	}
}
