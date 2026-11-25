//go:build !wasm

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
package axonvm

import (
	"bytes"
	"testing"
)

func runResponseWritePage(t *testing.T, src string) string {
	t.Helper()
	c := NewASPCompiler(src)
	if err := c.Compile(); err != nil {
		t.Fatalf("compile: %v", err)
	}
	vm := NewVMFromCompiler(c)
	host := NewMockHost()
	var out bytes.Buffer
	host.SetOutput(&out)
	vm.SetHost(host)
	if err := vm.Run(); err != nil {
		t.Fatalf("run: %v", err)
	}
	host.Response().Flush()
	return out.String()
}

// Response.Write(Null) called from VBScript with a parenthesised argument goes
// through the native member call instead of the flat write chain. IIS writes
// nothing for Null in both forms; the engine wrote the JScript text "null"
// once a JScript root environment existed, which happens on any page.
func TestVBScriptResponseWriteParenthesisedNullWritesNothing(t *testing.T) {
	got := runResponseWritePage(t, `<%
v = Null
Response.Write "[a:" : Response.Write v : Response.Write "]"
Response.Write "[b:" : Response.Write(v) : Response.Write "]"
Response.Write "[c:" : Response.Write(Null) : Response.Write "]"
Response.Write "[d:" : Call Response.Write(v) : Response.Write "]"
%>`)
	if want := "[a:][b:][c:][d:]"; got != want {
		t.Fatalf("expected %q, got %q", want, got)
	}
}

// Same on a page that also has a server-side JScript block.
func TestVBScriptResponseWriteParenthesisedNullWithJScriptBlock(t *testing.T) {
	got := runResponseWritePage(t, `<script language="javascript" runat="server">function helperNoop() { return 1; }</script><%
v = Null
Response.Write "[" : Response.Write(v) : Response.Write "]"
%>`)
	if want := "[]"; got != want {
		t.Fatalf("expected %q, got %q", want, got)
	}
}

// JScript keeps its own rule: Null renders as "null" when JScript writes it,
// both through the direct write statement and through a member call.
func TestJScriptResponseWriteNullStillWritesNull(t *testing.T) {
	got := runResponseWritePage(t, `<%@ Language="JScript" %><%
var v = null;
Response.Write(v);
var r = Response;
r.Write(v);
%>`)
	if want := "nullnull"; got != want {
		t.Fatalf("expected %q, got %q", want, got)
	}
}
