package axonvm

import (
	"bytes"
	"testing"
)

// TestCookieGitHubSubKeyEnum reproduces the exact ASP test scenario from
// www/tests/test_cookie_github.asp Test 3, where For Each is used on
// Request.Cookies(cn) when the cookie has sub-keys.
func TestCookieGitHubSubKeyEnum(t *testing.T) {
	source := `<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%Option Explicit%>
<%
' Create a main cookie named "user" containing multiple sub-keys
Response.Cookies("user")("firstname") = "John"
Response.Cookies("user")("lastname") = "Smith"
Response.Cookies("user")("country") = "Norway"
Response.Cookies("user")("age") = "25"

' Simulate the browser sending the cookie back on next request
' by directly populating Request.Cookies
Dim cn, kc, outStr
outStr = ""
For Each cn In Request.Cookies
    If Request.Cookies(cn).HasKeys Then
        For Each kc In Request.Cookies(cn)
            outStr = outStr & "[" & cn & "] " & kc & "=" & Request.Cookies(cn)(kc) & "|"
        Next
    End If
Next
Response.Write outStr
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()

	// Simulate browser sending the cookie with sub-keys
	host.Request().Cookies.AddCookie("user", "firstname=John&lastname=Smith&country=Norway&age=25")

	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	actual := output.String()
	// Keys sorted: age, country, firstname, lastname
	const want = "[user] age=25|[user] country=Norway|[user] firstname=John|[user] lastname=Smith|"
	if actual != want {
		t.Fatalf("unexpected output:\n  got:  %q\n  want: %q", actual, want)
	}
	t.Logf("Output: %q", actual)
}

// TestJScriptCookieLanguagePersistence verifies that JScript can read a
// previously-set language cookie via Request.Cookies("g3pix_lang") and
// correctly use String() for type conversion. This reproduces the exact
// G3PIX-WEBSITE language detection flow.
// Regression test: ensureJSRootEnv() must be initialized so that the
// JScript String() builtin (type conversion) is used instead of the
// VBScript String(n, char) builtin.
func TestJScriptCookieLanguagePersistence(t *testing.T) {
	source := `<%@Language="JavaScript"%>
<%
var rawCookie = String(Request.Cookies("g3pix_lang") || "");
Response.Write("rawCookie=" + rawCookie + "|");
Response.Write("len=" + rawCookie.length + "|");

var cookieLang = "";
if (rawCookie.length > 0) {
    cookieLang = rawCookie;
}
Response.Write("cookieLang=" + cookieLang + "|");

// Simulate T.detectLanguage
if (cookieLang === "pt" || cookieLang === "en") {
    Response.Write("lang=" + cookieLang);
} else {
    Response.Write("lang=fallback");
}
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()

	// Simulate browser sending the language cookie
	host.Request().Cookies.AddCookie("g3pix_lang", "pt")

	var output bytes.Buffer
	host.SetOutput(&output)
	host.Response().SetBuffer(false)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	actual := output.String()
	const want = "rawCookie=pt|len=2|cookieLang=pt|lang=pt"
	if actual != want {
		t.Fatalf("unexpected output:\n  got:  %q\n  want: %q", actual, want)
	}
	t.Logf("Output: %q", actual)
}

// TestJScriptCookieLanguageFallback verifies that when no language cookie
// is set, the JScript fallback path works correctly (empty cookie → domain
// detection → default language).
func TestJScriptCookieLanguageFallback(t *testing.T) {
	source := `<%@Language="JavaScript"%>
<%
var rawCookie = String(Request.Cookies("g3pix_lang") || "");
Response.Write("rawCookie=[" + rawCookie + "]|");
Response.Write("empty=" + (rawCookie === "" ? "true" : "false") + "|");

var cookieLang = "";
if (rawCookie.length > 0) {
    cookieLang = rawCookie;
}
Response.Write("cookieSet=" + (cookieLang.length === 0 ? "false" : "true"));
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()

	// No cookie set — simulate first visit

	var output bytes.Buffer
	host.SetOutput(&output)
	host.Response().SetBuffer(false)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	actual := output.String()
	const want = "rawCookie=[]|empty=true|cookieSet=false"
	if actual != want {
		t.Fatalf("unexpected output:\n  got:  %q\n  want: %q", actual, want)
	}
	t.Logf("Output: %q", actual)
}
