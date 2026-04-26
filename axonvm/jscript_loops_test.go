/*
 * AxonASP Server
 * Copyright (C) 2026 G3pix Ltda. All rights reserved.
 *
 * Developed by Lucas Guimarães - G3pix Ltda
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

import "testing"

func TestJScriptLoopOnlyCases(t *testing.T) {
	testCases := []struct {
		name   string
		source string
		want   string
	}{
		{
			name: "while loop",
			source: `<script runat="server" language="JScript">` +
				`var i = 0;` +
				`var out = "";` +
				`while (i < 3) { out += i + ","; i++; }` +
				`Response.Write(out);` +
				`</script>`,
			want: "0,1,2,",
		},
		{
			name: "do while loop",
			source: `<script runat="server" language="JScript">` +
				`var i = 0;` +
				`var out = "";` +
				`do { out += i + ","; i++; } while (i < 2);` +
				`Response.Write(out);` +
				`</script>`,
			want: "0,1,",
		},
		{
			name: "for loop",
			source: `<script runat="server" language="JScript">` +
				`var out = "";` +
				`for (var i = 0; i < 3; i++) { out += i + ","; }` +
				`Response.Write(out);` +
				`</script>`,
			want: "0,1,2,",
		},
		{
			name: "for break",
			source: `<script runat="server" language="JScript">` +
				`var out = "";` +
				`for (var i = 0; i < 5; i++) { if (i === 3) { break; } out += i + ","; }` +
				`Response.Write(out);` +
				`</script>`,
			want: "0,1,2,",
		},
		{
			name: "for continue",
			source: `<script runat="server" language="JScript">` +
				`var out = "";` +
				`for (var i = 0; i < 5; i++) { if (i === 2) { continue; } out += i + ","; }` +
				`Response.Write(out);` +
				`</script>`,
			want: "0,1,3,4,",
		},
		{
			name: "switch inside loop",
			source: `<script runat="server" language="JScript">` +
				`var out = "";` +
				`for (var i = 0; i < 4; i++) {` +
				`  switch (i) {` +
				`  case 0: out += "a"; break;` +
				`  case 1: continue;` +
				`  case 2: out += "c"; break;` +
				`  default: out += "d";` +
				`  }` +
				`}` +
				`Response.Write(out);` +
				`</script>`,
			want: "acd",
		},
		{
			name: "for in loop",
			source: `<script runat="server" language="JScript">` +
				`var o = {};` +
				`o.b = 2; o.a = 1; o.c = 3;` +
				`var out = "";` +
				`for (var key in o) { if (key === "b") { continue; } out += key; if (key === "c") { break; } }` +
				`Response.Write(out);` +
				`</script>`,
			want: "ac",
		},
		{
			name: "percent block for loop",
			source: `<%@ Language="JScript" %><%` +
				`var out = "";` +
				`for (var i = 0; i < 3; i++) { out += i; }` +
				`Response.Write(out);` +
				`%>`,
			want: "012",
		},
	}

	for _, tc := range testCases {
		t.Run(tc.name, func(t *testing.T) {
			out := runASPSourceForTest(t, tc.source)
			if out != tc.want {
				t.Fatalf("unexpected output: got %q want %q", out, tc.want)
			}
		})
	}
}

func TestJScriptForInLoopStateResetsAcrossFunctionCalls(t *testing.T) {
	source := `<script runat="server" language="JScript">` +
		`function iterateOnce(obj) {` +
		`  var out = "";` +
		`  for (var key in obj) { out += key; break; }` +
		`  return out;` +
		`}` +
		`var o = {};` +
		`o.a = 1; o.b = 2; o.c = 3;` +
		`Response.Write(iterateOnce(o) + "|" + iterateOnce(o));` +
		`</script>`

	out := runASPSourceForTest(t, source)
	if out != "a|a" {
		t.Fatalf("unexpected repeated for-in output: %q", out)
	}
}

func TestJScriptForInLoopInsideFunctionSingleCall(t *testing.T) {
	source := `<script runat="server" language="JScript">` +
		`function iterateOnce(obj) {` +
		`  var out = "";` +
		`  for (var key in obj) { out += key; break; }` +
		`  return out;` +
		`}` +
		`var o = {};` +
		`o.a = 1; o.b = 2; o.c = 3;` +
		`Response.Write(iterateOnce(o));` +
		`</script>`

	out := runASPSourceForTest(t, source)
	if out != "a" {
		t.Fatalf("unexpected single-call for-in output: %q", out)
	}
}

func TestJScriptForInLoopInsideFunctionWithoutBreak(t *testing.T) {
	source := `<script runat="server" language="JScript">` +
		`function iterateAll(obj) {` +
		`  var out = "";` +
		`  for (var key in obj) { out += key; }` +
		`  return out;` +
		`}` +
		`var o = {};` +
		`o.a = 1; o.b = 2; o.c = 3;` +
		`Response.Write(iterateAll(o));` +
		`</script>`

	out := runASPSourceForTest(t, source)
	if out != "abc" {
		t.Fatalf("unexpected full for-in output: %q", out)
	}
}

func TestJScriptLoopStressTerminates(t *testing.T) {
	source := `<script runat="server" language="JScript">` +
		`var i = 0;` +
		`while (i < 5000) { i++; }` +
		`Response.Write(i);` +
		`</script>`

	out := runASPSourceForTest(t, source)
	if out != "5000" {
		t.Fatalf("unexpected stress-loop output: %q", out)
	}
}

func TestJScriptObjectArgumentAccessInsideFunction(t *testing.T) {
	source := `<script runat="server" language="JScript">` +
		`function inspect(obj) { return typeof obj + "|" + obj.a + "|" + obj.b; }` +
		`var o = {};` +
		`o.a = 1; o.b = 2;` +
		`Response.Write(inspect(o));` +
		`</script>`

	out := runASPSourceForTest(t, source)
	if out != "object|1|2" {
		t.Fatalf("unexpected object-argument output: %q", out)
	}
}

func TestJScriptForInLoopInsideFunctionCanReturnFromBody(t *testing.T) {
	source := `<script runat="server" language="JScript">` +
		`function firstKey(obj) {` +
		`  for (var key in obj) { return "hit:" + key; }` +
		`  return "none";` +
		`}` +
		`var o = {};` +
		`o.a = 1; o.b = 2;` +
		`Response.Write(firstKey(o));` +
		`</script>`

	out := runASPSourceForTest(t, source)
	if out != "hit:a" {
		t.Fatalf("unexpected for-in body return output: %q", out)
	}
}
