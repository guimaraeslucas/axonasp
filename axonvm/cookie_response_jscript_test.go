package axonvm

import (
	"bytes"
	"testing"
)

// TestJScriptResponseCookieSet verifies that Response.Cookies("name") = value
// works correctly from JScript — cookie is set in response and flushed to output.
func TestJScriptResponseCookieSet(t *testing.T) {
	source := `<script runat="server" language="JScript">
Response.Cookies("g3pix_lang") = "pt";
Response.Cookies("g3pix_lang").Expires = new Date(Date.now() + 365 * 24 * 60 * 60 * 1000);
Response.Cookies("g3pix_lang").Path = "/";
Response.Write("ok");
</script>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	host.Response().SetBuffer(false)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	if output.String() != "ok" {
		t.Errorf("expected output 'ok', got %q", output.String())
	}

	// Verify cookie is in response
	cookieVal := host.Response().GetCookieValue("g3pix_lang")
	if cookieVal != "pt" {
		t.Errorf("expected cookie 'g3pix_lang' = 'pt', got %q", cookieVal)
	}
	t.Logf("Cookie value: %q", cookieVal)
	t.Logf("Output: %q", output.String())
}

// TestJScriptResponseCookieSetAspMode reproduces the exact default.asp
// cookie setting flow using <%@Language="JavaScript"%> directive (ASP mode).
func TestJScriptResponseCookieSetAspMode(t *testing.T) {
	source := `<%@Language="JavaScript"%>
<%
var queryLang = "pt";
Response.Cookies("g3pix_lang") = queryLang;
Response.Cookies("g3pix_lang").Expires = new Date(Date.now() + 365 * 24 * 60 * 60 * 1000);
Response.Cookies("g3pix_lang").Path = "/";
Response.Write("ok");
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	host.Response().SetBuffer(false)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	cookieVal := host.Response().GetCookieValue("g3pix_lang")
	if cookieVal != "pt" {
		t.Errorf("expected cookie 'g3pix_lang' = 'pt', got %q", cookieVal)
	}
	if output.String() != "ok" {
		t.Errorf("expected output 'ok', got %q", output.String())
	}
	t.Logf("Cookie value: %q, Output: %q", cookieVal, output.String())
}
