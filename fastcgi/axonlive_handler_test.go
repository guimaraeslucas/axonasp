//go:build !lib_g3axonlive_disabled

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

package main

import (
	"bytes"
	"encoding/json"
	"io"
	"maps"
	"net"
	"net/http"
	"net/http/fcgi"
	"net/http/httptest"
	"os"
	"path/filepath"
	"testing"
	"time"

	"g3pix.com.br/axonasp/axonvm"
)

// TestNormalizeAndAuthorizeG3ALSessionID verifies fetch payload session IDs are bound to ASPSESSIONID.
func TestNormalizeAndAuthorizeG3ALSessionID(t *testing.T) {
	tests := []struct {
		name           string
		payloadSession string
		authSession    string
		expectedID     string
		expectedOK     bool
	}{
		{name: "missing authenticated session", payloadSession: "ABC", authSession: "", expectedID: "", expectedOK: false},
		{name: "payload omitted", payloadSession: "", authSession: "AUTH123", expectedID: "AUTH123", expectedOK: true},
		{name: "payload matches authenticated", payloadSession: "AUTH123", authSession: "AUTH123", expectedID: "AUTH123", expectedOK: true},
		{name: "payload case-insensitive match", payloadSession: "auth123", authSession: "AUTH123", expectedID: "AUTH123", expectedOK: true},
		{name: "payload mismatch", payloadSession: "OTHER999", authSession: "AUTH123", expectedID: "", expectedOK: false},
	}

	for _, tc := range tests {
		t.Run(tc.name, func(t *testing.T) {
			gotID, gotOK := normalizeAndAuthorizeG3ALSessionID(tc.payloadSession, tc.authSession)
			if gotOK != tc.expectedOK {
				t.Fatalf("expected ok=%v, got %v", tc.expectedOK, gotOK)
			}
			if gotID != tc.expectedID {
				t.Fatalf("expected session ID %q, got %q", tc.expectedID, gotID)
			}
		})
	}
}

// TestAuthenticatedSessionIDFromRequest verifies ASPSESSIONID extraction from request cookies.
func TestAuthenticatedSessionIDFromRequest(t *testing.T) {
	req := httptest.NewRequest(http.MethodPost, "http://example.local/g3al", nil)
	req.AddCookie(&http.Cookie{Name: "ASPSESSIONID", Value: "  AUTH555  "})

	got := authenticatedSessionIDFromRequest(req)
	if got != "AUTH555" {
		t.Fatalf("expected trimmed session ID AUTH555, got %q", got)
	}
}

// TestResolveFCGIG3ALScriptURL verifies page lookup uses the authenticated session mapping.
func TestResolveFCGIG3ALScriptURL(t *testing.T) {
	authSession := "RESOLVE123"
	scriptURL := "/axonlive/counter.asp"
	axonvm.G3ALRegisterPage(authSession, scriptURL)

	resolved := resolveFCGIG3ALScriptURL(authSession)
	if resolved != scriptURL {
		t.Fatalf("expected script URL %q, got %q", scriptURL, resolved)
	}
}

// TestHandleFCGIG3AxonLiveFetch_DocumentRoot verifies that the handler correctly
// uses the DOCUMENT_ROOT from the FastCGI environment when resolving the page path.
func TestHandleFCGIG3AxonLiveFetch_DocumentRoot(t *testing.T) {
	// Setup: Create a temporary document root outside the default RootDir
	tmpDir, err := os.MkdirTemp("", "axonasp-test-root")
	if err != nil {
		t.Fatalf("failed to create temp dir: %v", err)
	}
	defer os.RemoveAll(tmpDir)

	// Create a dummy ASP file in the temp root
	aspFile := filepath.Join(tmpDir, "test.asp")
	if err := os.WriteFile(aspFile, []byte("<% Response.Write \"OK\" %>"), 0644); err != nil {
		t.Fatalf("failed to write test asp file: %v", err)
	}

	// Save original RootDir and restore it later
	originalRootDir := RootDir
	defer func() { RootDir = originalRootDir }()
	// Set RootDir to a non-existent path to ensure it's NOT used if DOCUMENT_ROOT is present.
	RootDir = filepath.Join(os.TempDir(), "non-existent-axonasp-root")

	// Register the page for a session
	sessionID := "PATH-TEST-123"
	scriptURL := "/test.asp"
	axonvm.G3ALRegisterPage(sessionID, scriptURL)

	// Mock the G3AxonLive event
	event := FCGIAxonLiveEvent{
		SessionID:   sessionID,
		ComponentID: "comp1",
		EventName:   "click",
	}
	eventData, _ := json.Marshal(event)

	// Create request with DOCUMENT_ROOT param using the helper from main_test.go
	// We need to inject the body into the FCGI stream.
	req := newFCGIRequestWithBody(t, http.MethodPost, "/g3al/", map[string]string{
		"DOCUMENT_ROOT":     tmpDir,
		"HTTP_X_G3AXONLIVE": "true",
		"CONTENT_TYPE":      "application/json",
	}, eventData)

	// Add the session cookie
	req.AddCookie(&http.Cookie{Name: "ASPSESSIONID", Value: sessionID})

	w := httptest.NewRecorder()

	// Execute the handler.
	// It will likely fail at executeASP because the VM environment isn't fully set up,
	// but we're testing that it DOESN'T fail with 403/404 before that.
	handleFCGIG3AxonLiveFetch(w, req)

	// If the fix works, it should have resolved the path correctly.
	// If it used RootDir, it would have returned 403 (outside root) or 404.
	if w.Code == http.StatusForbidden || w.Code == http.StatusNotFound {
		t.Fatalf("Handler failed with status %d; likely failed to resolve path using DOCUMENT_ROOT", w.Code)
	}
}

// newFCGIRequestWithBody is a variation of newFCGIRequest from main_test.go that supports a POST body.
func newFCGIRequestWithBody(t *testing.T, method, urlPath string, cgiEnv map[string]string, body []byte) *http.Request {
	t.Helper()

	params := map[string]string{
		"REQUEST_METHOD":  method,
		"SERVER_PROTOCOL": "HTTP/1.1",
		"SERVER_NAME":     "localhost",
		"SERVER_PORT":     "80",
		"SCRIPT_NAME":     urlPath,
		"REQUEST_URI":     urlPath,
	}
	maps.Copy(params, cgiEnv)

	var paramContent []byte
	for k, v := range params {
		paramContent = append(paramContent, fcgiBuildNameValue(k, v)...)
	}

	var msg bytes.Buffer
	msg.Write(fcgiBuildRecord(1, 1, []byte{0, 1, 0, 0, 0, 0, 0, 0})) // BEGIN_REQUEST
	msg.Write(fcgiBuildRecord(4, 1, paramContent))                   // PARAMS
	msg.Write(fcgiBuildRecord(4, 1, nil))                            // PARAMS end
	if len(body) > 0 {
		msg.Write(fcgiBuildRecord(5, 1, body)) // STDIN
	}
	msg.Write(fcgiBuildRecord(5, 1, nil)) // STDIN end

	reqCh := make(chan *http.Request, 1)
	serverConn, clientConn := net.Pipe()
	ln := &chanListener{ch: make(chan net.Conn, 1), addr: serverConn.LocalAddr()}
	ln.ch <- serverConn

	handler := http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		select {
		case reqCh <- r:
		default:
		}
		w.WriteHeader(http.StatusOK)
	})

	go func() { _ = fcgi.Serve(ln, handler) }()

	go func() {
		_, _ = io.Copy(clientConn, &msg)
		_ = clientConn.Close()
	}()

	select {
	case req := <-reqCh:
		return req
	case <-time.After(2 * time.Second):
		t.Fatal("timeout waiting for fcgi.Serve to capture request")
		return nil
	}
}
