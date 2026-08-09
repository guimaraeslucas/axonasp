package axonvm

import (
	"testing"
)

// TestJScriptActiveXObjectBinding verifies that new ActiveXObject and ActiveXObject()
// correctly instantiate native COM/ActiveX objects with method and property resolution.
func TestJScriptActiveXObjectBinding(t *testing.T) {
	source := `<script runat="server" language="JScript">` +
		`var fso = new ActiveXObject("Scripting.FileSystemObject");` +
		`var exists = fso.FileExists("nonexistent_file_123456789.txt");` +
		`var folder = fso.GetFolder(".");` +
		`var fso2 = ActiveXObject("Scripting.FileSystemObject");` +
		`Response.Write(typeof fso);` +
		`Response.Write("|");` +
		`Response.Write(exists);` +
		`Response.Write("|");` +
		`Response.Write(typeof folder);` +
		`Response.Write("|");` +
		`Response.Write(typeof folder.Files);` +
		`Response.Write("|");` +
		`Response.Write(typeof fso2);` +
		`</script>`
	out := runASPSourceForTest(t, source)
	if out != "object|False|object|object|object" {
		t.Fatalf("unexpected ActiveXObject output: %q", out)
	}
}

// TestJScriptRegExpStaticProperties verifies state mutation of RegExp.$1...$9 and context aliases post-execution.
func TestJScriptRegExpStaticProperties(t *testing.T) {
	source := `<script runat="server" language="JScript">` +
		`var re = /([a-z]+)\s+([0-9]+)/;` +
		`var res = re.exec("prefix hello 123 suffix");` +
		`Response.Write(RegExp.$1 + "|");` +
		`Response.Write(RegExp.$2 + "|");` +
		`Response.Write(RegExp.$3 + "|");` +
		`Response.Write(RegExp.input + "|");` +
		`Response.Write(RegExp.$_ + "|");` +
		`Response.Write(RegExp.lastMatch + "|");` +
		`Response.Write(RegExp["$&"] + "|");` +
		`Response.Write(RegExp.lastParen + "|");` +
		`Response.Write(RegExp["$+"] + "|");` +
		`Response.Write(RegExp.leftContext + "|");` +
		"Response.Write(RegExp[\"$`\"] + \"|\");" +
		`Response.Write(RegExp.rightContext + "|");` +
		`Response.Write(RegExp["$'"]);` +
		`</script>`
	out := runASPSourceForTest(t, source)
	expected := "hello|123||prefix hello 123 suffix|prefix hello 123 suffix|hello 123|hello 123|123|123|prefix |prefix | suffix| suffix"
	if out != expected {
		t.Fatalf("unexpected RegExp static properties output:\ngot:  %q\nwant: %q", out, expected)
	}
}

// TestJScriptFunctionCalleeAndCaller verifies typeof arguments.callee and typeof Function.caller evaluate to "function".
func TestJScriptFunctionCalleeAndCaller(t *testing.T) {
	source := `<script runat="server" language="JScript">` +
		`function testCallee() { return typeof arguments.callee; }` +
		`function inner() { return typeof Function.caller + "|" + typeof arguments.callee.caller; }` +
		`function outer() { var res = inner(); return res; }` +
		`Response.Write(testCallee() + "|");` +
		`Response.Write(outer());` +
		`</script>`
	out := runASPSourceForTest(t, source)
	expected := "function|function|function"
	if out != expected {
		t.Fatalf("unexpected function properties output:\ngot:  %q\nwant: %q", out, expected)
	}
}
