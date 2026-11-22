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

// MSXML dataType "bin.base64"/"bin.hex" with nodeTypedValue: the classic ASP
// Base64/hex encoding idiom (it returned "").
func TestMSXMLNodeTypedValueBinary(t *testing.T) {
	cases := map[string]string{
		// bytes -> base64 text
		`<% Set x = CreateObject("Msxml2.DOMDocument.3.0") : Set n = x.CreateElement("b") : n.dataType = "bin.base64" : n.nodeTypedValue = ChrB(49) & ChrB(50) & ChrB(51) & ChrB(95) & ChrB(49) : Response.Write n.dataType & "|" & n.text %>`: "bin.base64|MTIzXzE=",
		// base64 text -> bytes
		`<% Set x = CreateObject("Msxml2.DOMDocument.3.0") : Set n = x.CreateElement("b") : n.dataType = "bin.base64" : n.text = "MTIzXzE=" : v = n.nodeTypedValue : Response.Write LenB(v) & "|" & AscB(v) %>`: "5|49",
		// hex both ways
		`<% Set x = CreateObject("Msxml2.DOMDocument.3.0") : Set n = x.CreateElement("h") : n.dataType = "bin.hex" : n.nodeTypedValue = ChrB(65) & ChrB(255) : Response.Write n.text & "|" & AscB(MidB(n.nodeTypedValue, 2, 1)) %>`: "41ff|255",
		// untyped node: nodeTypedValue is the text
		`<% Set x = CreateObject("Msxml2.DOMDocument.3.0") : Set n = x.CreateElement("p") : n.text = "hi" : Response.Write n.nodeTypedValue %>`: "hi",
	}
	for src, want := range cases {
		if got := runVBSAndGetOutput(t, src); got != want {
			t.Errorf("%s => %q, want %q", src, got, want)
		}
	}
}
