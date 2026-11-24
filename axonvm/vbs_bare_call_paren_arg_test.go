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

// TestVBSBareCallParenthesisedArgumentContinues checks that in a call statement
// whose first argument starts with "(", the operators after ")" stay part of
// that argument: "Foo (a) & b" passes "(a) & b", as IIS does.
func TestVBSBareCallParenthesisedArgumentContinues(t *testing.T) {
	src := `<%
Sub Show(s)
  Response.Write "[" & s & "]"
End Sub
Sub Two(a, b)
  Response.Write "[" & a & "," & b & "]"
End Sub
Sub Bump(n)
  n = n + 1
End Sub
Class Box
  Public Sub Show(s)
    Response.Write "{" & s & "}"
  End Sub
  Public Sub Run()
    Me.Show ("m") & "n"
  End Sub
End Class
Show ("a") & "b"
Show ("a") & "b" & "c"
Show (1) + 2 * 3
Show ("a")
Two ("a") & "b", "c"
Two ("a"), "b"
Response.Write (2 + 3) * 2
Response.Write "|"
Response.Write (1 = 1) & "|"
x = 5
Bump (x)
Response.Write x & "|"
Dim arr(2)
arr(1) = "z"
Response.Write arr(1) & "|"
Set b = New Box
b.Show ("p") & "q"
b.Run
Set xmlDoc = CreateObject("MSXML2.DOMDocument")
xmlDoc.loadXML("<?xml version=""1.0""?><Types>" & "<Type n=""A""/><Type n=""C""/>") & "</Types>"
Response.Write "|" & xmlDoc.selectNodes("/Types/Type").length
%>`
	got := runVBSAndGetOutput(t, src)
	want := "[ab][abc][7][a][ab,c][a,b]10|True|5|z|{pq}{mn}|2"
	if got != want {
		t.Fatalf("got %q, want %q", got, want)
	}
}
