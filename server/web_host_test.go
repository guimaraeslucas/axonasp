/*
 * AxonASP Server
 * Copyright (C) 2026 G3pix Ltda. All rights reserved.
 *
 * Developed by Lucas GuimarÃ£es - G3pix Ltda
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
package main

import (
	"bytes"
	"io"
	"net"
	"net/http"
	"net/http/httptest"
	"path/filepath"
	"strings"
	"testing"

	"g3pix.com.br/axonasp/axonvm"
	"g3pix.com.br/axonasp/axonvm/asp"
)

type countingReadCloser struct {
	reader io.Reader
	reads  int
}

func newCountingReadCloser(payload string) *countingReadCloser {
	return &countingReadCloser{reader: strings.NewReader(payload)}
}

func (c *countingReadCloser) Read(p []byte) (int, error) {
	n, err := c.reader.Read(p)
	if n > 0 {
		c.reads++
	}
	return n, err
}

func (c *countingReadCloser) Close() error { return nil }

// sessionCookieFromResponse returns the ASPSESSIONID cookie from a response recorder.
func sessionCookieFromResponse(t *testing.T, recorder *httptest.ResponseRecorder) *http.Cookie {
	t.Helper()

	for _, cookie := range recorder.Result().Cookies() {
		if cookie.Name == sessionCookieName {
			return cookie
		}
	}

	t.Fatalf("expected %s cookie in response", sessionCookieName)
	return nil
}

// TestNewWebHostSetsSessionCookie verifies NewWebHost emits ASPSESSIONID for new sessions.
func TestNewWebHostSetsSessionCookie(t *testing.T) {
	asp.SetSessionStorageDir(t.TempDir())
	defer asp.SetSessionStorageDir(filepath.Join("temp", "session"))

	req := httptest.NewRequest(http.MethodGet, "http://example.local/default.asp", nil)
	recorder := httptest.NewRecorder()
	host := NewWebHost(recorder, req)

	if host.Session() == nil || host.Session().ID == "" {
		t.Fatalf("expected host session with non-empty ID")
	}

	cookie := sessionCookieFromResponse(t, recorder)
	if cookie.Value != host.Session().ID {
		t.Fatalf("expected cookie value %q, got %q", host.Session().ID, cookie.Value)
	}
}

// TestNewWebHostReusesExistingSession verifies cookie-bound sessions are loaded and reused.
func TestNewWebHostReusesExistingSession(t *testing.T) {
	asp.SetSessionStorageDir(t.TempDir())
	defer asp.SetSessionStorageDir(filepath.Join("temp", "session"))

	existing, err := asp.CreateSession()
	if err != nil {
		t.Fatalf("failed to create baseline session: %v", err)
	}

	req := httptest.NewRequest(http.MethodGet, "http://example.local/default.asp", nil)
	req.AddCookie(&http.Cookie{Name: sessionCookieName, Value: existing.ID})
	recorder := httptest.NewRecorder()
	host := NewWebHost(recorder, req)

	if host.Session() == nil {
		t.Fatalf("expected host session")
	}
	if host.Session().ID != existing.ID {
		t.Fatalf("expected reused session ID %q, got %q", existing.ID, host.Session().ID)
	}

	cookie := sessionCookieFromResponse(t, recorder)
	if cookie.Value != existing.ID {
		t.Fatalf("expected response cookie value %q, got %q", existing.ID, cookie.Value)
	}
}

// TestWebHostClassLifecycleParity verifies class lifecycle semantics execute consistently in HTTP host path.
func TestWebHostClassLifecycleParity(t *testing.T) {
	source := `<%
Class Widget
	Private Sub Class_Initialize()
		Response.Write "I"
	End Sub
	Private Sub Class_Terminate()
		Response.Write "T"
	End Sub
End Class

Dim w
Set w = New Widget
Response.Write "K"
%>`

	request := httptest.NewRequest(http.MethodGet, "http://example.local/default.asp", nil)
	recorder := httptest.NewRecorder()
	host := NewWebHost(recorder, request)

	compiler := axonvm.NewASPCompiler(source)
	compiler.SetSourceName("/default.asp")
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := axonvm.NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	vm.SetHost(host)
	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	if recorder.Body.String() != "IKT" {
		t.Fatalf("expected IKT output, got %q", recorder.Body.String())
	}
}

// TestNewWebHostPopulatesRequestCollections verifies form fields and common server variables.
func TestNewWebHostPopulatesRequestCollections(t *testing.T) {
	body := strings.NewReader("txtName=Ada&txtAge=30")
	req := httptest.NewRequest(http.MethodPost, "http://example.local:8080/tests/claude.asp?color=crimson", body)
	req.Header.Set("Content-Type", "application/x-www-form-urlencoded")
	req.Header.Set("Accept-Language", "pt-BR,pt;q=0.9")
	req.RemoteAddr = net.JoinHostPort("203.0.113.9", "54321")
	recorder := httptest.NewRecorder()
	host := NewWebHost(recorder, req)

	if got := host.Request().Form.Get("txtAge"); got != "30" {
		t.Fatalf("expected form txtAge=30, got %q", got)
	}
	if got := host.Request().QueryString.Get("color"); got != "crimson" {
		t.Fatalf("expected query color=crimson, got %q", got)
	}
	if got := host.Request().ServerVars.Get("SERVER_NAME"); got != "example.local" {
		t.Fatalf("expected SERVER_NAME example.local, got %q", got)
	}
	if got := host.Request().ServerVars.Get("SERVER_PORT"); got != "8080" {
		t.Fatalf("expected SERVER_PORT 8080, got %q", got)
	}
	if got := host.Request().ServerVars.Get("SCRIPT_NAME"); got != "/tests/claude.asp" {
		t.Fatalf("expected SCRIPT_NAME /tests/claude.asp, got %q", got)
	}
	if got := host.Request().ServerVars.Get("HTTP_ACCEPT_LANGUAGE"); got != "pt-BR,pt;q=0.9" {
		t.Fatalf("expected HTTP_ACCEPT_LANGUAGE header, got %q", got)
	}
	if got := host.Request().ServerVars.Get("REMOTE_ADDR"); got != "203.0.113.9" {
		t.Fatalf("expected REMOTE_ADDR without port, got %q", got)
	}
}

// TestNewWebHostDoesNotEagerReadGETBody verifies GET requests do not preload the request body.
func TestNewWebHostDoesNotEagerReadGETBody(t *testing.T) {
	body := newCountingReadCloser("alpha=1")
	req := httptest.NewRequest(http.MethodGet, "http://example.local/default.asp", body)
	req.ContentLength = int64(len("alpha=1"))
	rec := httptest.NewRecorder()
	host := NewWebHost(rec, req)

	if body.reads != 0 {
		t.Fatalf("expected no body reads during host creation for GET, got %d", body.reads)
	}
	if got := host.Request().TotalBytes(); got != int64(len("alpha=1")) {
		t.Fatalf("expected TotalBytes from content length, got %d", got)
	}
}

// TestNewWebHostLoadsBodyOnFormAccess verifies POST body/form parsing is deferred until Form is accessed.
func TestNewWebHostLoadsBodyOnFormAccess(t *testing.T) {
	body := newCountingReadCloser("name=ada&age=30")
	req := httptest.NewRequest(http.MethodPost, "http://example.local/default.asp", body)
	req.Header.Set("Content-Type", "application/x-www-form-urlencoded")
	rec := httptest.NewRecorder()
	host := NewWebHost(rec, req)

	if body.reads != 0 {
		t.Fatalf("expected no body reads before Form access, got %d", body.reads)
	}
	if got := host.Request().Form.Get("age"); got != "30" {
		t.Fatalf("expected deferred form parse to return age=30, got %q", got)
	}
	if body.reads == 0 {
		t.Fatalf("expected body reads after Form access")
	}
}

// TestNewWebHostLoadsBodyOnBinaryRead verifies BinaryRead triggers deferred request body loading.
func TestNewWebHostLoadsBodyOnBinaryRead(t *testing.T) {
	body := newCountingReadCloser("abcdef")
	req := httptest.NewRequest(http.MethodPost, "http://example.local/default.asp", body)
	rec := httptest.NewRecorder()
	host := NewWebHost(rec, req)

	if body.reads != 0 {
		t.Fatalf("expected no body reads before BinaryRead, got %d", body.reads)
	}
	if got := string(host.Request().BinaryRead(3)); got != "abc" {
		t.Fatalf("expected BinaryRead to return abc, got %q", got)
	}
	if body.reads == 0 {
		t.Fatalf("expected body reads after BinaryRead")
	}
}

// TestWebHostBinaryResponsePreservesHeadersAndBytes verifies HTTP responses keep binary bytes intact and do not append a charset to image content types.
func TestWebHostBinaryResponsePreservesHeadersAndBytes(t *testing.T) {
	source := `<%
Response.Buffer = True
Response.ContentType = "image/bmp"
Response.BinaryWrite ChrB(&H42)
Response.BinaryWrite ChrB(&H4D)
Response.BinaryWrite ChrB(&H80)
Response.Flush
%>`

	req := httptest.NewRequest(http.MethodGet, "http://example.local/tests/captcha.asp", nil)
	recorder := httptest.NewRecorder()
	host := NewWebHost(recorder, req)

	compiler := axonvm.NewASPCompiler(source)
	compiler.SetSourceName("/tests/captcha.asp")
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := axonvm.NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	vm.SetHost(host)
	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	if got := recorder.Header().Get("Content-Type"); got != "image/bmp" {
		t.Fatalf("expected image/bmp content type, got %q", got)
	}
	if !bytes.Equal(recorder.Body.Bytes(), []byte{'B', 'M', 0x80}) {
		t.Fatalf("unexpected binary body: %v", recorder.Body.Bytes())
	}
}
