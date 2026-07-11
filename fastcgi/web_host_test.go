//go:build !wasm

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
	"io"
	"net"
	"net/http"
	"net/http/httptest"
	"os"
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

// fastCGISessionCookieFromResponse returns the ASPSESSIONID cookie from a response recorder.
func fastCGISessionCookieFromResponse(t *testing.T, recorder *httptest.ResponseRecorder) *http.Cookie {
	t.Helper()

	for _, cookie := range recorder.Result().Cookies() {
		if cookie.Name == fastCGISessionCookieName {
			return cookie
		}
	}

	t.Fatalf("expected %s cookie in response", fastCGISessionCookieName)
	return nil
}

// TestNewFastCGIHostSetsSessionCookie verifies NewFastCGIHost emits ASPSESSIONID for new sessions.
func TestNewFastCGIHostSetsSessionCookie(t *testing.T) {
	asp.SetSessionStorageDir(t.TempDir())
	defer asp.SetSessionStorageDir(filepath.Join("temp", "session"))

	req := httptest.NewRequest(http.MethodGet, "http://example.local/default.asp", nil)
	recorder := httptest.NewRecorder()
	host := NewFastCGIHost(recorder, req)

	if host.Session() == nil || host.Session().ID == "" {
		t.Fatalf("expected host session with non-empty ID")
	}

	cookie := fastCGISessionCookieFromResponse(t, recorder)
	if cookie.Value != host.Session().ID {
		t.Fatalf("expected cookie value %q, got %q", host.Session().ID, cookie.Value)
	}
}

// TestNewFastCGIHostReusesExistingSession verifies cookie-bound sessions are loaded and reused.
func TestNewFastCGIHostReusesExistingSession(t *testing.T) {
	asp.SetSessionStorageDir(t.TempDir())
	defer asp.SetSessionStorageDir(filepath.Join("temp", "session"))

	existing, err := asp.CreateSession()
	if err != nil {
		t.Fatalf("failed to create baseline session: %v", err)
	}

	req := httptest.NewRequest(http.MethodGet, "http://example.local/default.asp", nil)
	req.AddCookie(&http.Cookie{Name: fastCGISessionCookieName, Value: existing.ID})
	recorder := httptest.NewRecorder()
	host := NewFastCGIHost(recorder, req)

	if host.Session() == nil {
		t.Fatalf("expected host session")
	}
	if host.Session().ID != existing.ID {
		t.Fatalf("expected reused session ID %q, got %q", existing.ID, host.Session().ID)
	}

	cookie := fastCGISessionCookieFromResponse(t, recorder)
	if cookie.Value != existing.ID {
		t.Fatalf("expected response cookie value %q, got %q", existing.ID, cookie.Value)
	}
}

// TestFastCGIHostPersistSessionKeepsSingleSessionCookie verifies session persistence rewrites ASPSESSIONID instead of duplicating it.
func TestFastCGIHostPersistSessionKeepsSingleSessionCookie(t *testing.T) {
	asp.SetSessionStorageDir(t.TempDir())
	defer asp.SetSessionStorageDir(filepath.Join("temp", "session"))

	req := httptest.NewRequest(http.MethodGet, "http://example.local/default.asp", nil)
	recorder := httptest.NewRecorder()
	host := NewFastCGIHost(recorder, req)

	host.PersistSession()

	count := 0
	for _, cookie := range recorder.Result().Cookies() {
		if cookie.Name == fastCGISessionCookieName {
			count++
			if cookie.Value != host.Session().ID {
				t.Fatalf("expected response cookie value %q, got %q", host.Session().ID, cookie.Value)
			}
		}
	}

	if count != 1 {
		t.Fatalf("expected 1 %s cookie, got %d", fastCGISessionCookieName, count)
	}
}

// TestFastCGIHostClassLifecycleParity verifies class lifecycle semantics execute consistently in FastCGI host path.
func TestFastCGIHostClassLifecycleParity(t *testing.T) {
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
	host := NewFastCGIHost(recorder, request)

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

// TestNewFastCGIHostPopulatesRequestCollections verifies form fields and common server variables.
func TestNewFastCGIHostPopulatesRequestCollections(t *testing.T) {
	body := strings.NewReader("txtName=Ada&txtAge=30")
	req := httptest.NewRequest(http.MethodPost, "http://example.local:8080/tests/claude.asp?color=crimson", body)
	req.Header.Set("Content-Type", "application/x-www-form-urlencoded")
	req.Header.Set("Accept-Language", "pt-BR,pt;q=0.9")
	req.RemoteAddr = net.JoinHostPort("203.0.113.9", "54321")
	recorder := httptest.NewRecorder()
	host := NewFastCGIHost(recorder, req)

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

// TestNewFastCGIHostDoesNotEagerReadGETBody verifies GET requests do not preload body bytes.
func TestNewFastCGIHostDoesNotEagerReadGETBody(t *testing.T) {
	body := newCountingReadCloser("alpha=1")
	req := httptest.NewRequest(http.MethodGet, "http://example.local/default.asp", body)
	req.ContentLength = int64(len("alpha=1"))
	rec := httptest.NewRecorder()
	host := NewFastCGIHost(rec, req)

	if body.reads != 0 {
		t.Fatalf("expected no body reads during host creation for GET, got %d", body.reads)
	}
	if got := host.Request().TotalBytes(); got != int64(len("alpha=1")) {
		t.Fatalf("expected TotalBytes from content length, got %d", got)
	}
}

// TestNewFastCGIHostLoadsBodyOnFormAccess verifies POST body/form parsing is deferred until Form is accessed.
func TestNewFastCGIHostLoadsBodyOnFormAccess(t *testing.T) {
	body := newCountingReadCloser("name=ada&age=30")
	req := httptest.NewRequest(http.MethodPost, "http://example.local/default.asp", body)
	req.Header.Set("Content-Type", "application/x-www-form-urlencoded")
	rec := httptest.NewRecorder()
	host := NewFastCGIHost(rec, req)

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

// TestNewFastCGIHostLoadsBodyOnBinaryRead verifies BinaryRead triggers deferred body loading.
func TestNewFastCGIHostLoadsBodyOnBinaryRead(t *testing.T) {
	body := newCountingReadCloser("abcdef")
	req := httptest.NewRequest(http.MethodPost, "http://example.local/default.asp", body)
	rec := httptest.NewRecorder()
	host := NewFastCGIHost(rec, req)

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

// TestNewFastCGIHostUsesDocumentRootForMapPath verifies FastCGI per-request
// DOCUMENT_ROOT drives Server.MapPath resolution instead of global RootDir.
func TestNewFastCGIHostUsesDocumentRootForMapPath(t *testing.T) {
	tmpDir := t.TempDir()
	docRoot := filepath.Join(tmpDir, "site-a")
	if err := os.MkdirAll(filepath.Join(docRoot, "manual", "assets"), 0755); err != nil {
		t.Fatalf("failed to create doc root layout: %v", err)
	}

	originalRootDir := RootDir
	defer func() { RootDir = originalRootDir }()
	RootDir = filepath.Join(tmpDir, "configured-root")

	req := newFCGIRequest(t, http.MethodGet, "/manual/default.asp", map[string]string{
		"DOCUMENT_ROOT": docRoot,
	})
	rec := httptest.NewRecorder()
	host := NewFastCGIHost(rec, req)

	resolved := host.Server().MapPath("assets/file.js")
	want := filepath.Join(docRoot, "manual", "assets", "file.js")
	wantAbs, err := filepath.Abs(want)
	if err != nil {
		t.Fatalf("failed to resolve expected absolute path: %v", err)
	}
	if resolved != wantAbs {
		t.Fatalf("expected MapPath relative resolution %q, got %q", wantAbs, resolved)
	}
}

// TestNewFastCGIHostFallsBackToRootDirWithoutDocumentRoot verifies backward
// compatibility when reverse proxy does not provide DOCUMENT_ROOT.
func TestNewFastCGIHostFallsBackToRootDirWithoutDocumentRoot(t *testing.T) {
	tmpDir := t.TempDir()
	configuredRoot := filepath.Join(tmpDir, "configured-root")
	if err := os.MkdirAll(filepath.Join(configuredRoot, "manual", "assets"), 0755); err != nil {
		t.Fatalf("failed to create configured root layout: %v", err)
	}

	originalRootDir := RootDir
	defer func() { RootDir = originalRootDir }()
	RootDir = configuredRoot

	req := httptest.NewRequest(http.MethodGet, "http://example.local/manual/default.asp", nil)
	rec := httptest.NewRecorder()
	host := NewFastCGIHost(rec, req)

	resolved := host.Server().MapPath("assets/file.js")
	want := filepath.Join(configuredRoot, "manual", "assets", "file.js")
	wantAbs, err := filepath.Abs(want)
	if err != nil {
		t.Fatalf("failed to resolve expected absolute path: %v", err)
	}
	if resolved != wantAbs {
		t.Fatalf("expected fallback MapPath relative resolution %q, got %q", wantAbs, resolved)
	}
}

// TestNewFastCGIHostSessionOnStartEndedDoesNotBlankPage verifies that when
// Session_OnStart calls Response.End(), the subsequent page execution still
// produces output instead of silently discarding everything (blank page).
// This reproduces the "first request always blank" regression where
// Response.ended leaked from Session_OnStart into page execution.
func TestNewFastCGIHostSessionOnStartEndedDoesNotBlankPage(t *testing.T) {
	asp.SetSessionStorageDir(t.TempDir())
	defer asp.SetSessionStorageDir(filepath.Join("temp", "session"))

	tmpDir := t.TempDir()

	// Write a global.asa with Session_OnStart that calls Response.End()
	globalASAContent := `<script runat="server" language="VBScript">
Sub Session_OnStart
	Response.End
End Sub
</script>`
	globalASAPath := filepath.Join(tmpDir, "global.asa")
	if err := os.WriteFile(globalASAPath, []byte(globalASAContent), 0644); err != nil {
		t.Fatalf("failed to write global.asa: %v", err)
	}

	// Reset global ASA singleton for this test
	axonvm.ResetGlobalASA()

	// Load and compile the test global.asa
	app := asp.NewApplication()
	if err := axonvm.GetGlobalASA().LoadAndCompile(tmpDir, app); err != nil {
		t.Fatalf("failed to load/compile global.asa: %v", err)
	}

	// First request: creates a new session, Session_OnStart fires
	req := httptest.NewRequest(http.MethodGet, "http://example.local/default.asp", nil)
	recorder := httptest.NewRecorder()
	host := NewFastCGIHost(recorder, req)

	// Verify session was created
	if host.Session() == nil || host.Session().ID == "" {
		t.Fatalf("expected non-empty session ID")
	}

	// Execute a simple ASP page
	source := `<% Response.Write "HELLO_WORLD" %>`
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

	body := recorder.Body.String()
	if body != "HELLO_WORLD" {
		t.Fatalf("expected body HELLO_WORLD, got %q (blank page regression: Response.ended leaked from Session_OnStart)", body)
	}

	// Verify the session cookie was set despite Response.End in Session_OnStart
	cookie := fastCGISessionCookieFromResponse(t, recorder)
	if cookie.Value != host.Session().ID {
		t.Fatalf("expected cookie value %q, got %q", host.Session().ID, cookie.Value)
	}
}

// TestNewFastCGIHostSessionOnStartRedirectDoesNotBlankPage verifies that
// when Session_OnStart calls Response.Redirect, the subsequent page still
// produces output instead of silently discarding everything.
func TestNewFastCGIHostSessionOnStartRedirectDoesNotBlankPage(t *testing.T) {
	asp.SetSessionStorageDir(t.TempDir())
	defer asp.SetSessionStorageDir(filepath.Join("temp", "session"))

	tmpDir := t.TempDir()

	// Write a global.asa with Session_OnStart that calls Response.Redirect
	globalASAContent := `<script runat="server" language="VBScript">
Sub Session_OnStart
	Response.Redirect "/login.asp"
End Sub
</script>`
	globalASAPath := filepath.Join(tmpDir, "global.asa")
	if err := os.WriteFile(globalASAPath, []byte(globalASAContent), 0644); err != nil {
		t.Fatalf("failed to write global.asa: %v", err)
	}

	// Reset global ASA singleton for this test
	axonvm.ResetGlobalASA()

	// Load and compile the test global.asa
	app := asp.NewApplication()
	if err := axonvm.GetGlobalASA().LoadAndCompile(tmpDir, app); err != nil {
		t.Fatalf("failed to load/compile global.asa: %v", err)
	}

	// First request: creates a new session, Session_OnStart fires
	req := httptest.NewRequest(http.MethodGet, "http://example.local/default.asp", nil)
	recorder := httptest.NewRecorder()
	host := NewFastCGIHost(recorder, req)

	// Execute a simple ASP page
	source := `<% Response.Write "PAGE_OK" %>`
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

	body := recorder.Body.String()
	if body != "PAGE_OK" {
		t.Fatalf("expected body PAGE_OK, got %q (Response.Redirect in Session_OnStart leaked ended flag)", body)
	}
}

// TestNewFastCGIHostSubsequentRequestDoesNotRunSessionOnStart verifies that
// the second request (with existing ASPSESSIONID cookie) does not execute
// Session_OnStart, preserving page output without the ended-reset path.
func TestNewFastCGIHostSubsequentRequestDoesNotRunSessionOnStart(t *testing.T) {
	asp.SetSessionStorageDir(t.TempDir())
	defer asp.SetSessionStorageDir(filepath.Join("temp", "session"))

	tmpDir := t.TempDir()

	// Write a global.asa with a panic-inducing Session_OnStart to prove it's skipped
	globalASAContent := `<script runat="server" language="VBScript">
Sub Session_OnStart
	' This MUST NOT execute on the second request
	Response.Write "SESSION_START_SHOULD_NOT_APPEAR"
End Sub
</script>`
	globalASAPath := filepath.Join(tmpDir, "global.asa")
	if err := os.WriteFile(globalASAPath, []byte(globalASAContent), 0644); err != nil {
		t.Fatalf("failed to write global.asa: %v", err)
	}

	// Reset global ASA singleton for this test
	axonvm.ResetGlobalASA()

	// Load and compile the test global.asa
	app := asp.NewApplication()
	if err := axonvm.GetGlobalASA().LoadAndCompile(tmpDir, app); err != nil {
		t.Fatalf("failed to load/compile global.asa: %v", err)
	}

	// Create a pre-existing session to simulate second request
	existing, err := asp.CreateSession()
	if err != nil {
		t.Fatalf("failed to create baseline session: %v", err)
	}

	// Second request: has ASPSESSIONID cookie, session is NOT new
	req := httptest.NewRequest(http.MethodGet, "http://example.local/default.asp", nil)
	req.AddCookie(&http.Cookie{Name: fastCGISessionCookieName, Value: existing.ID})
	recorder := httptest.NewRecorder()
	host := NewFastCGIHost(recorder, req)

	// Execute a simple ASP page
	source := `<% Response.Write "SECOND_REQUEST_OK" %>`
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

	body := recorder.Body.String()
	if body != "SECOND_REQUEST_OK" {
		t.Fatalf("expected body SECOND_REQUEST_OK, got %q", body)
	}
}
