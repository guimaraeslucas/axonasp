package vbscript

import (
	"strings"
	"testing"
)

func TestScanASPSourceRegions(t *testing.T) {
	text := "<%@ Language=JScript %>\r\n😀<% var x=1; %><% = x %><!-- #include file=\"part.asp\" -->" + `<script language="JavaScript" runat="server">var y="</script>";</script>`
	regions, err := ScanASP(text)
	if err != nil {
		t.Fatal(err)
	}
	var kinds []string
	for _, r := range regions {
		kinds = append(kinds, r.Kind)
		if r.Kind == "expression" && text[r.CodeStart:r.CodeEnd] != " x " {
			t.Fatalf("bad expression span: %+v", r)
		}
		if r.Kind == "include" && r.Path != "part.asp" {
			t.Fatal(r)
		}
	}
	if strings.Join(kinds, ",") != "directive,static,script,expression,include,script" {
		t.Fatal(kinds)
	}
}

func TestScanASPRejectsTruncatedBlocks(t *testing.T) {
	for _, text := range []string{"<% var x=1;", `<script runat="server" language="JScript">var x=1;`} {
		if _, err := ScanASP(text); err == nil {
			t.Fatalf("accepted %q", text)
		}
	}
}

func TestParseASPIncludeCommentMatchesPreprocessor(t *testing.T) {
	for _, tc := range []struct {
		comment, path  string
		virtual, valid bool
	}{
		{`<!-- #include file="part.asp" -->`, "part.asp", false, true},
		{`<!--#INCLUDE VIRTUAL='/part.asp'-->`, "/part.asp", true, true},
		{`<!-- #include file=part.asp -->`, "", false, false},
		{`<!-- ordinary comment -->`, "", false, false},
	} {
		path, virtual, ok := ParseASPIncludeComment(tc.comment)
		if path != tc.path || virtual != tc.virtual || ok != tc.valid {
			t.Fatalf("%q: %q %v %v", tc.comment, path, virtual, ok)
		}
	}
}
