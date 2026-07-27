/*
 * AxonASP Server
 * Copyright (C) 2026 G3pix Ltda. All rights reserved.
 */
package axonvm

import (
	"strings"
	"testing"
)

func TestJScriptCollectionCompatibility(t *testing.T) {
	host := NewMockHost()
	host.Request().QueryString.Add("category", "books")
	host.Request().QueryString.AddValues("tag", []string{"scifi", "thriller"})

	source := `<script runat="server" language="JScript">
		var objRQ = Request.QueryString;

		// 1. Coerce collection to string
		Response.Write("1:" + String(objRQ) + "\n");

		// 2. Count on collection item
		Response.Write("2:" + objRQ("category").Count + "\n");

		// 3. Item on collection item
		Response.Write("3:" + objRQ("tag").Item() + "\n");
		Response.Write("4:" + objRQ("tag").Item(1) + "\n");
		Response.Write("5:" + objRQ("tag").Item(2) + "\n");

		// 4. Index out of range throws
		try {
			objRQ("tag").Item(0);
			Response.Write("6:fail\n");
		} catch(err) {
			Response.Write("6:ok:" + err.number + ":" + err.description + "\n");
		}

		try {
			objRQ("fish").Item(1);
			Response.Write("7:fail\n");
		} catch(err) {
			Response.Write("7:ok:" + err.number + ":" + err.description + "\n");
		}

		// 5. Unsupported method throws
		try {
			objRQ("tag").Key();
			Response.Write("8:fail\n");
		} catch(err) {
			Response.Write("8:ok:" + err.number + ":" + err.description + "\n");
		}

		// 6. Enumerator yields keys
		var e = new Enumerator(objRQ);
		var keys = [];
		for(; !e.atEnd(); e.moveNext()) {
			keys.push(e.item());
		}
		Response.Write("9:" + keys.join(",") + "\n");

		// 7. Nonexistent collection item primitive coercion returns undefined in JS
		Response.Write("10:" + String(objRQ("fish")) + "\n");
	</script>`

	out := runASPSourceForTestWithHost(t, source, host)
	lines := strings.Split(strings.TrimSpace(out), "\n")

	expected := []string{
		"1:category=books&tag=scifi&tag=thriller",
		"2:1",
		"3:scifi, thriller",
		"4:scifi",
		"5:thriller",
		"6:ok:-2147467259:007~ASP 0105~Index out of range~An array index is out of range.",
		"7:ok:-2147467259:007~ASP 0105~Index out of range~An array index is out of range.",
		"8:ok:-2146827850:Object doesn't support this property or method",
		"9:category,tag",
		"10:undefined",
	}

	if len(lines) != len(expected) {
		t.Fatalf("expected %d output lines, got %d:\n%s", len(expected), len(lines), out)
	}

	for i, exp := range expected {
		if lines[i] != exp {
			t.Errorf("line %d: expected %q, got %q", i+1, exp, lines[i])
		}
	}
}

func TestNormalizeJScriptCollectionAssignmentsRegexEscapedQuotes(t *testing.T) {
	input := `
function escapeAttr(s) {
  return (s + '').replace(/\\/g, "\\\\").replace(/\'/g, "&#39;").replace(/\"/g, "&quot;");
}
var sItems = escapeAttr("hello 'world'");
Response.Cookies("SelectedItem") = sItems;
`
	expected := `
function escapeAttr(s) {
  return (s + '').replace(/\\/g, "\\\\").replace(/\'/g, "&#39;").replace(/\"/g, "&quot;");
}
var sItems = escapeAttr("hello 'world'");
Response.Cookies("SelectedItem", sItems);
`
	got := normalizeJScriptCollectionAssignments(input)
	if got != expected {
		t.Errorf("normalizeJScriptCollectionAssignments failed.\nExpected:\n%s\nGot:\n%s", expected, got)
	}
}

func TestJScriptCollectionAssignmentWithEscapedQuoteRegex(t *testing.T) {
	source := `<%@LANGUAGE="JScript"%>
<%
function escapeAttr(s) {
  return (s + '').replace(/\\/g, "\\\\").replace(/\'/g, "&#39;").replace(/\"/g, "&quot;");
}

var sItems = escapeAttr("hello 'world'");

Response.Cookies("SelectedItem") = sItems;
Response.Write("done");
%>`

	out := runASPSourceForTest(t, source)
	if strings.TrimSpace(out) != "done" {
		t.Fatalf("expected output 'done', got %q", out)
	}
}
