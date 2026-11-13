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

import (
	"bytes"
	"net/http"
	"net/http/httptest"
	"testing"
)

// TestVBScriptDictionaryCookiePersistenceHTTPRequest verifies that VBScript
// Response.Cookies("MGSsession")("key2") = val preserves existing incoming
// subkeys received from the HTTP request, supports HasKeys and For Each, and
// resolves VTNativeObject values like Request.Form("item").
func TestVBScriptDictionaryCookiePersistenceHTTPRequest(t *testing.T) {
	source := `<%
Response.Cookies("MGSsession").Path = "/app"
Response.Cookies("MGSsession")("key2") = "val2"
Response.Cookies("MGSsession")("fromform") = Request.Form("item")
Response.Cookies("MGSsession")("key1") = "updated1"

Dim keys
keys = ""
For Each k In Response.Cookies("MGSsession")
	If keys <> "" Then keys = keys & ","
	keys = keys & k
Next

Response.Write Response.Cookies("MGSsession").HasKeys & "|" & _
	Response.Cookies("MGSsession")("key1") & "|" & _
	Response.Cookies("MGSsession")("key2") & "|" & _
	Response.Cookies("MGSsession")("fromform") & "|" & _
	keys & "|" & _
	Response.Cookies("MGSsession").Path
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	req := httptest.NewRequest(http.MethodPost, "http://example.local/", nil)
	req.Header.Set("Cookie", "MGSsession=key1=val1&existing=keepme")

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	host.Response().SetRequest(req)
	host.Request().Cookies.AddCookie("MGSsession", "key1=val1&existing=keepme")
	host.Request().Form.Add("item", "alice")
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	wantOut := "True|updated1|val2|alice|key1,existing,key2,fromform|/app"
	if got := output.String(); got != wantOut {
		t.Fatalf("unexpected VBScript output: got %q want %q", got, wantOut)
	}

	wantCookie := "key1=updated1&existing=keepme&key2=val2&fromform=alice"
	if got := host.Response().GetCookieValue("MGSsession"); got != wantCookie {
		t.Fatalf("unexpected cookie value: got %q want %q", got, wantCookie)
	}
}

// TestVBScriptDictionaryCookiePersistenceMockHostFallback verifies that even when
// host.Response().SetRequest is nil (such as MockHost), incoming Request.Cookies
// subkeys are preserved when assigning a single subkey on Response.Cookies.
func TestVBScriptDictionaryCookiePersistenceMockHostFallback(t *testing.T) {
	source := `<%
Response.Cookies("MGSsession")("key2") = "val2"
Response.Write Response.Cookies("MGSsession")
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	host.Request().Cookies.AddCookie("MGSsession", "key1=val1")
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	if got := host.Response().GetCookieValue("MGSsession"); got != "key1=val1&key2=val2" {
		t.Fatalf("expected MockHost fallback subkey persistence, got %q", got)
	}
	if got := output.String(); got != "key1=val1&key2=val2" {
		t.Fatalf("unexpected output: got %q", got)
	}
}

// TestVBScriptDictionaryCookieExplicitClearBeforeSubKey verifies that
// Response.Cookies("MGSsession") = "" explicitly drops incoming request subkeys
// before subsequent subkey writes.
func TestVBScriptDictionaryCookieExplicitClearBeforeSubKey(t *testing.T) {
	source := `<%
Response.Cookies("MGSsession") = ""
Response.Cookies("MGSsession")("fresh") = "only"
Response.Write Response.Cookies("MGSsession")
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	req := httptest.NewRequest(http.MethodGet, "http://example.local/", nil)
	req.Header.Set("Cookie", "MGSsession=key1=val1&key2=val2")

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	host.Response().SetRequest(req)
	host.Request().Cookies.AddCookie("MGSsession", "key1=val1&key2=val2")
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	if got := host.Response().GetCookieValue("MGSsession"); got != "fresh=only" {
		t.Fatalf("expected explicit reset before subkey write, got %q", got)
	}
}

// TestJScriptDictionaryCookiePersistence verifies that JScript ASP pages preserve
// incoming dictionary cookie subkeys when writing individual subkeys.
func TestJScriptDictionaryCookiePersistence(t *testing.T) {
	source := `<%@Language="JavaScript"%>
<%
Response.Cookies("MGSsession")("key2") = "val2";
Response.Write(Response.Cookies("MGSsession").HasKeys + "|" + Response.Cookies("MGSsession")("key1") + "|" + Response.Cookies("MGSsession")("key2"));
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	req := httptest.NewRequest(http.MethodGet, "http://example.local/", nil)
	req.Header.Set("Cookie", "MGSsession=key1=val1")

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	host.Response().SetRequest(req)
	host.Request().Cookies.AddCookie("MGSsession", "key1=val1")
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	if got := output.String(); got != "true|val1|val2" {
		t.Fatalf("unexpected JScript output: got %q want %q", got, "true|val1|val2")
	}
	if got := host.Response().GetCookieValue("MGSsession"); got != "key1=val1&key2=val2" {
		t.Fatalf("unexpected JScript cookie value: got %q", got)
	}
}
