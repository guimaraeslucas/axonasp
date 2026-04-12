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
	"io"
	"net"
	"net/http"
	"net/http/fcgi"
	"net/http/httptest"
	"os"
	"path/filepath"
	"strings"
	"sync"
	"testing"
	"time"
)

func TestParseFastCGIListenEndpoint(t *testing.T) {
	tests := []struct {
		name      string
		input     string
		wantNet   string
		wantAddr  string
		wantError bool
	}{
		{
			name:     "numeric tcp port",
			input:    "9000",
			wantNet:  "tcp",
			wantAddr: "127.0.0.1:9000",
		},
		{
			name:     "host port",
			input:    "0.0.0.0:9100",
			wantNet:  "tcp",
			wantAddr: "0.0.0.0:9100",
		},
		{
			name:     "unix socket endpoint",
			input:    "unix:/tmp/axonasp.sock",
			wantNet:  "unix",
			wantAddr: "/tmp/axonasp.sock",
		},
		{
			name:      "invalid tcp port",
			input:     "70000",
			wantError: true,
		},
		{
			name:      "missing unix socket path",
			input:     "unix:",
			wantError: true,
		},
		{
			name:      "invalid endpoint",
			input:     "localhost",
			wantError: true,
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			gotNet, gotAddr, err := parseFastCGIListenEndpoint(tt.input)
			if tt.wantError {
				if err == nil {
					t.Fatalf("expected error, got network=%q address=%q", gotNet, gotAddr)
				}
				return
			}

			if err != nil {
				t.Fatalf("unexpected error: %v", err)
			}
			if gotNet != tt.wantNet {
				t.Fatalf("network mismatch: got %q want %q", gotNet, tt.wantNet)
			}
			if gotAddr != tt.wantAddr {
				t.Fatalf("address mismatch: got %q want %q", gotAddr, tt.wantAddr)
			}
		})
	}
}

// chanListener is a single-use net.Listener backed by a channel of pre-created
// net.Conns, allowing fcgi.Serve to be driven by an in-memory net.Pipe pair.
type chanListener struct {
	ch   chan net.Conn
	addr net.Addr
	once sync.Once
}

func (l *chanListener) Accept() (net.Conn, error) {
	conn, ok := <-l.ch
	if !ok {
		return nil, net.ErrClosed
	}
	return conn, nil
}
func (l *chanListener) Close() error   { l.once.Do(func() { close(l.ch) }); return nil }
func (l *chanListener) Addr() net.Addr { return l.addr }

// fcgiBuildRecord constructs a raw FastCGI record (version=1).
func fcgiBuildRecord(typ byte, reqID uint16, content []byte) []byte {
	hdr := []byte{
		1, typ,
		byte(reqID >> 8), byte(reqID),
		byte(len(content) >> 8), byte(len(content)),
		0, 0,
	}
	return append(hdr, content...)
}

// fcgiBuildNameValue encodes a single CGI name=value pair per the FastCGI spec.
func fcgiBuildNameValue(name, value string) []byte {
	encLen := func(n int) []byte {
		if n <= 127 {
			return []byte{byte(n)}
		}
		return []byte{byte(n>>24) | 0x80, byte(n >> 16), byte(n >> 8), byte(n)}
	}
	var buf []byte
	buf = append(buf, encLen(len(name))...)
	buf = append(buf, encLen(len(value))...)
	buf = append(buf, []byte(name)...)
	buf = append(buf, []byte(value)...)
	return buf
}

// newFCGIRequest creates an *http.Request with proper CGI environment variables
// in its context by routing the request through a real fcgi.Serve instance over
// an in-memory net.Pipe connection. This is required because fcgi.ProcessEnv reads
// from an unexported context key that is only populated by fcgi.Serve itself.
// DOCUMENT_ROOT, SCRIPT_FILENAME, and other non-standard CGI params can be passed
// via cgiEnv and will be accessible via getFastCGIParam / fcgi.ProcessEnv.
// NOTE: SCRIPT_NAME and HTTP_* vars are consumed by cgi.RequestFromMap and will NOT
// appear in fcgi.ProcessEnv — SCRIPT_NAME becomes r.URL.Path automatically.
func newFCGIRequest(t *testing.T, method, urlPath string, cgiEnv map[string]string) *http.Request {
	t.Helper()

	// Build the full FastCGI params map.
	params := map[string]string{
		"REQUEST_METHOD":  method,
		"SERVER_PROTOCOL": "HTTP/1.1",
		"SERVER_NAME":     "localhost",
		"SERVER_PORT":     "80",
		"SCRIPT_NAME":     urlPath,
		"REQUEST_URI":     urlPath,
	}
	for k, v := range cgiEnv {
		params[k] = v
	}

	// Encode all params as FastCGI name-value pairs.
	var paramContent []byte
	for k, v := range params {
		paramContent = append(paramContent, fcgiBuildNameValue(k, v)...)
	}

	// Build the full FastCGI message: BEGIN_REQUEST → PARAMS → PARAMS(empty) → STDIN(empty)
	var msg bytes.Buffer
	msg.Write(fcgiBuildRecord(1, 1, []byte{0, 1, 0, 0, 0, 0, 0, 0})) // BEGIN_REQUEST, role=responder
	msg.Write(fcgiBuildRecord(4, 1, paramContent))                   // PARAMS
	msg.Write(fcgiBuildRecord(4, 1, nil))                            // PARAMS end
	msg.Write(fcgiBuildRecord(5, 1, nil))                            // STDIN end

	reqCh := make(chan *http.Request, 1)
	serverConn, clientConn := net.Pipe()
	ln := &chanListener{ch: make(chan net.Conn, 1), addr: serverConn.LocalAddr()}
	ln.ch <- serverConn

	// Handler captures the request; returns 200 immediately so fcgi.Serve can complete.
	handler := http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		select {
		case reqCh <- r:
		default:
		}
		w.WriteHeader(http.StatusOK)
	})

	go func() { _ = fcgi.Serve(ln, handler) }()
	go func() {
		defer clientConn.Close()
		clientConn.Write(msg.Bytes())   //nolint:errcheck
		io.Copy(io.Discard, clientConn) //nolint:errcheck
	}()
	t.Cleanup(func() { ln.Close() })

	select {
	case req := <-reqCh:
		return req
	case <-time.After(5 * time.Second):
		t.Fatal("newFCGIRequest: timeout waiting for request from fcgi.Serve")
		return nil
	}
}

// TestGetFastCGIParam verifies that getFastCGIParam reads from the CGI environment
// context populated by fcgi.Serve (via fcgi.ProcessEnv). It also documents that
// SCRIPT_NAME is filtered out of ProcessEnv (it becomes r.URL.Path instead).
func TestGetFastCGIParam(t *testing.T) {
	tests := []struct {
		name        string
		cgiEnv      map[string]string // nil = use plain httptest.NewRequest (no FastCGI context)
		paramName   string
		expected    string
		description string
	}{
		{
			name:        "DOCUMENT_ROOT from real FastCGI request",
			cgiEnv:      map[string]string{"DOCUMENT_ROOT": "/var/www/site1"},
			paramName:   "DOCUMENT_ROOT",
			expected:    "/var/www/site1",
			description: "DOCUMENT_ROOT is stored in context by fcgi.Serve via fcgi.ProcessEnv",
		},
		{
			name:        "SCRIPT_FILENAME from real FastCGI request",
			cgiEnv:      map[string]string{"SCRIPT_FILENAME": "/var/www/site1/index.asp"},
			paramName:   "SCRIPT_FILENAME",
			expected:    "/var/www/site1/index.asp",
			description: "SCRIPT_FILENAME is stored in context by fcgi.Serve via fcgi.ProcessEnv",
		},
		{
			name:        "missing parameter returns empty",
			cgiEnv:      map[string]string{},
			paramName:   "DOCUMENT_ROOT",
			expected:    "",
			description: "Should return empty string when parameter is not in FastCGI env",
		},
		{
			name:        "SCRIPT_NAME is filtered from ProcessEnv",
			cgiEnv:      map[string]string{},
			paramName:   "SCRIPT_NAME",
			expected:    "",
			description: "SCRIPT_NAME is consumed by cgi.RequestFromMap (goes to r.URL.Path), excluded from ProcessEnv",
		},
		{
			name:        "plain request without FastCGI context",
			cgiEnv:      nil,
			paramName:   "DOCUMENT_ROOT",
			expected:    "",
			description: "Should return empty when request has no FastCGI context (httptest.NewRequest)",
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			var req *http.Request
			if tt.cgiEnv == nil {
				req = httptest.NewRequest("GET", "http://localhost/", nil)
			} else {
				req = newFCGIRequest(t, "GET", "/", tt.cgiEnv)
			}

			result := getFastCGIParam(req, tt.paramName)
			if result != tt.expected {
				t.Errorf("getFastCGIParam(%q) = %q, want %q\n%s", tt.paramName, result, tt.expected, tt.description)
			}
		})
	}
}

// TestFastCGIMiddlewareSetsPoweredByHeader verifies FastCGI responses emit X-Powered-By: AxonASP.
func TestFastCGIMiddlewareSetsPoweredByHeader(t *testing.T) {
	rec := httptest.NewRecorder()
	req := httptest.NewRequest(http.MethodGet, "http://example.local/", nil)
	handler := fastCGIMiddleware(func(w http.ResponseWriter, r *http.Request) {
		w.WriteHeader(http.StatusNoContent)
	})

	handler(rec, req)

	if got := rec.Header().Get("X-Powered-By"); got != "AxonASP" {
		t.Fatalf("expected X-Powered-By header AxonASP, got %q", got)
	}
}

// TestHandleRequestWithDocumentRoot verifies that handleRequest correctly
// resolves files using DOCUMENT_ROOT from FastCGI parameters.
func TestHandleRequestWithDocumentRoot(t *testing.T) {
	// Setup: Create temporary directories and test files
	tmpDir := t.TempDir()
	site1Dir := filepath.Join(tmpDir, "site1")
	site2Dir := filepath.Join(tmpDir, "site2")

	if err := os.MkdirAll(site1Dir, 0755); err != nil {
		t.Fatal(err)
	}
	if err := os.MkdirAll(site2Dir, 0755); err != nil {
		t.Fatal(err)
	}

	// Create test files in each site directory
	site1File := filepath.Join(site1Dir, "test.html")
	site2File := filepath.Join(site2Dir, "test.html")
	site1Content := "Site 1 Content"
	site2Content := "Site 2 Content"

	if err := os.WriteFile(site1File, []byte(site1Content), 0644); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(site2File, []byte(site2Content), 0644); err != nil {
		t.Fatal(err)
	}

	tests := []struct {
		name               string
		documentRoot       string
		url                string
		expectedStatusCode int
		expectedContent    string
		description        string
	}{
		{
			name:               "resolve from site1",
			documentRoot:       site1Dir,
			url:                "/test.html",
			expectedStatusCode: http.StatusOK,
			expectedContent:    site1Content,
			description:        "Should resolve test.html from site1 document root",
		},
		{
			name:               "same URL resolves different file per DOCUMENT_ROOT",
			documentRoot:       site2Dir,
			url:                "/test.html",
			expectedStatusCode: http.StatusOK,
			expectedContent:    site2Content,
			description:        "Same URL path returns different content based on DOCUMENT_ROOT",
		},
		{
			name:               "missing file returns 404",
			documentRoot:       site1Dir,
			url:                "/notfound.html",
			expectedStatusCode: http.StatusNotFound,
			expectedContent:    "",
			description:        "Should return 404 for missing files",
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			cgiEnv := map[string]string{}
			if tt.documentRoot != "" {
				cgiEnv["DOCUMENT_ROOT"] = tt.documentRoot
			}
			req := newFCGIRequest(t, "GET", tt.url, cgiEnv)

			w := httptest.NewRecorder()
			handleRequest(w, req)

			if w.Code != tt.expectedStatusCode {
				t.Errorf("Expected status %d, got %d\n%s", tt.expectedStatusCode, w.Code, tt.description)
			}

			if tt.expectedContent != "" {
				body, _ := io.ReadAll(w.Body)
				if !strings.Contains(string(body), tt.expectedContent) {
					t.Errorf("Expected response to contain %q, got %q\n%s", tt.expectedContent, string(body), tt.description)
				}
			}
		})
	}
}

// TestHandleRequestFallbackToRootDir verifies that handleRequest falls back to
// RootDir when DOCUMENT_ROOT is not provided (backward compatibility).
func TestHandleRequestFallbackToRootDir(t *testing.T) {
	// Setup: Create a test file in RootDir
	tmpDir := t.TempDir()
	testFile := filepath.Join(tmpDir, "test.html")
	testContent := "Test Content"

	if err := os.WriteFile(testFile, []byte(testContent), 0644); err != nil {
		t.Fatal(err)
	}

	// Save original RootDir
	originalRootDir := RootDir
	defer func() { RootDir = originalRootDir }()
	RootDir = tmpDir

	// Create request without DOCUMENT_ROOT header (backward compatibility)
	req := httptest.NewRequest("GET", "http://localhost/test.html", nil)

	w := httptest.NewRecorder()
	handleRequest(w, req)

	if w.Code != http.StatusOK {
		t.Errorf("Expected status 200, got %d (should fall back to RootDir)", w.Code)
	}

	body, _ := io.ReadAll(w.Body)
	if !strings.Contains(string(body), testContent) {
		t.Errorf("Expected response to contain %q, got %q", testContent, string(body))
	}
}

// TestDirectoryTraversalPrevention verifies that handleRequest prevents
// directory traversal attacks via URL path sequences with "..".
func TestDirectoryTraversalPrevention(t *testing.T) {
	tmpDir := t.TempDir()
	siteDir := filepath.Join(tmpDir, "site")
	parentFile := filepath.Join(tmpDir, "secret.txt")

	if err := os.MkdirAll(siteDir, 0755); err != nil {
		t.Fatal(err)
	}

	// Create a file outside the document root (simulating what attacker tries to access)
	if err := os.WriteFile(parentFile, []byte("Secret"), 0644); err != nil {
		t.Fatal(err)
	}

	// Override RootDir to be siteDir so no DOCUMENT_ROOT param is needed.
	originalRootDir := RootDir
	defer func() { RootDir = originalRootDir }()
	RootDir = siteDir

	// Craft a request whose URL.Path tries to escape via ".."
	// httptest.NewRequest preserves the raw path without normalizing ".."
	req := httptest.NewRequest("GET", "http://localhost/", nil)
	req.URL.Path = "/subdir/../../secret.txt" // bypass URL parser normalization

	w := httptest.NewRecorder()
	handleRequest(w, req)

	if w.Code != http.StatusForbidden {
		t.Errorf("Expected status 403 (Forbidden) for directory traversal, got %d", w.Code)
	}
}

// TestURLPathRouting verifies that r.URL.Path (set from FastCGI SCRIPT_NAME by
// cgi.RequestFromMap) controls which file is served when no DOCUMENT_ROOT is given.
func TestURLPathRouting(t *testing.T) {
	tmpDir := t.TempDir()
	testFile := filepath.Join(tmpDir, "test.asp")

	if err := os.WriteFile(testFile, []byte("test"), 0644); err != nil {
		t.Fatal(err)
	}

	originalRootDir := RootDir
	defer func() { RootDir = originalRootDir }()
	RootDir = tmpDir

	// Request where r.URL.Path = /test.asp; file exists, should return 200.
	req := httptest.NewRequest("GET", "http://localhost/test.asp", nil)

	w := httptest.NewRecorder()
	handleRequest(w, req)

	// Should resolve test.asp (from SCRIPT_NAME), not wrong.asp (from URL.Path)
	if w.Code != http.StatusOK {
		t.Errorf("Expected 200 OK when URL path resolves to an existing file, got %d", w.Code)
	}
}

// TestDefaultPages verifies that default pages (index.asp, default.asp, etc.)
// are correctly served when a directory is requested.
func TestDefaultPages(t *testing.T) {
	tmpDir := t.TempDir()

	// Create directory with default page
	subDir := filepath.Join(tmpDir, "subdir")
	if err := os.MkdirAll(subDir, 0755); err != nil {
		t.Fatal(err)
	}

	defaultPage := filepath.Join(subDir, "default.asp")
	if err := os.WriteFile(defaultPage, []byte("Default Page"), 0644); err != nil {
		t.Fatal(err)
	}

	originalRootDir := RootDir
	defer func() { RootDir = originalRootDir }()
	RootDir = tmpDir

	// Request directory without trailing slash
	req := httptest.NewRequest("GET", "http://localhost/subdir", nil)
	w := httptest.NewRecorder()
	handleRequest(w, req)

	// Should redirect to /subdir/ (with trailing slash)
	if w.Code != http.StatusMovedPermanently {
		t.Errorf("Expected 301 redirect for directory without slash, got %d", w.Code)
	}

	// Request directory with trailing slash
	req = httptest.NewRequest("GET", "http://localhost/subdir/", nil)
	w = httptest.NewRecorder()
	handleRequest(w, req)

	if w.Code != http.StatusOK {
		t.Errorf("Expected 200 OK for directory with default page, got %d", w.Code)
	}
}

// TestEmptyDocumentRootFallback verifies that empty DOCUMENT_ROOT is treated
// as not provided and falls back to RootDir.
func TestEmptyDocumentRootFallback(t *testing.T) {
	tmpDir := t.TempDir()
	testFile := filepath.Join(tmpDir, "test.html")

	if err := os.WriteFile(testFile, []byte("Content"), 0644); err != nil {
		t.Fatal(err)
	}

	originalRootDir := RootDir
	defer func() { RootDir = originalRootDir }()
	RootDir = tmpDir

	// Request with whitespace-only DOCUMENT_ROOT in CGI env; should fall back to RootDir.
	req := newFCGIRequest(t, "GET", "/test.html", map[string]string{
		"DOCUMENT_ROOT": "   ",
	})
	w := httptest.NewRecorder()
	handleRequest(w, req)

	if w.Code != http.StatusOK {
		t.Errorf("Expected fallback to RootDir when DOCUMENT_ROOT is empty, got status %d", w.Code)
	}
}

// TestRootSlashRequest verifies that "/" requests are handled correctly.
func TestRootSlashRequest(t *testing.T) {
	tmpDir := t.TempDir()

	// Create a default.asp in root
	defaultPage := filepath.Join(tmpDir, "default.asp")
	if err := os.WriteFile(defaultPage, []byte("Root"), 0644); err != nil {
		t.Fatal(err)
	}

	originalRootDir := RootDir
	defer func() { RootDir = originalRootDir }()
	RootDir = tmpDir

	req := httptest.NewRequest("GET", "http://localhost/", nil)
	w := httptest.NewRecorder()
	handleRequest(w, req)

	if w.Code != http.StatusOK {
		t.Errorf("Expected 200 OK for root directory request, got %d", w.Code)
	}
}

// TestMultipleVirtualHosts simulates serving two different document roots
// from a single FastCGI process (typical nginx + FastCGI setup).
func TestMultipleVirtualHosts(t *testing.T) {
	// Create two separate document roots
	tmpDir := t.TempDir()
	host1Dir := filepath.Join(tmpDir, "host1")
	host2Dir := filepath.Join(tmpDir, "host2")

	if err := os.MkdirAll(host1Dir, 0755); err != nil {
		t.Fatal(err)
	}
	if err := os.MkdirAll(host2Dir, 0755); err != nil {
		t.Fatal(err)
	}

	// Create different files in each host directory
	if err := os.WriteFile(filepath.Join(host1Dir, "site.html"), []byte("Host 1"), 0644); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(filepath.Join(host2Dir, "site.html"), []byte("Host 2"), 0644); err != nil {
		t.Fatal(err)
	}

	// Request to host1 — CGI env supplies its own DOCUMENT_ROOT
	req1 := newFCGIRequest(t, "GET", "/site.html", map[string]string{
		"DOCUMENT_ROOT": host1Dir,
	})
	w1 := httptest.NewRecorder()
	handleRequest(w1, req1)

	// Request to host2 — different DOCUMENT_ROOT in CGI env
	req2 := newFCGIRequest(t, "GET", "/site.html", map[string]string{
		"DOCUMENT_ROOT": host2Dir,
	})
	w2 := httptest.NewRecorder()
	handleRequest(w2, req2)

	// Both should succeed but with different content
	if w1.Code != http.StatusOK || w2.Code != http.StatusOK {
		t.Errorf("Expected both hosts to return 200 OK")
	}

	body1, _ := io.ReadAll(w1.Body)
	body2, _ := io.ReadAll(w2.Body)

	if !strings.Contains(string(body1), "Host 1") {
		t.Errorf("Host 1 should return its content")
	}
	if !strings.Contains(string(body2), "Host 2") {
		t.Errorf("Host 2 should return its content")
	}
}
