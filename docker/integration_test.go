/*
 * AxonASP Server - Docker Integration Tests
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

// Package docker contains HTTP integration tests that verify AxonASP works
// correctly inside a Docker container (or any running instance).
//
// The tests target a live server whose base URL is read from the
// TEST_SERVER_URL environment variable (default: http://localhost:4050).
//
// Run against an already-started container:
//
//	TEST_SERVER_URL=http://localhost:4050 go test ./docker/... -v -run TestDocker
package docker

import (
	"fmt"
	"io"
	"net/http"
	"os"
	"strings"
	"testing"
	"time"
)

// baseURL returns the server URL under test.
func baseURL() string {
	if u := os.Getenv("TEST_SERVER_URL"); u != "" {
		return strings.TrimRight(u, "/")
	}
	return "http://localhost:4050"
}

// newClient returns an HTTP client with a sensible timeout.
func newClient() *http.Client {
	return &http.Client{Timeout: 15 * time.Second}
}

// get performs a GET request and returns the status code and response body.
func get(t *testing.T, path string) (int, string) {
	t.Helper()
	url := baseURL() + path
	resp, err := newClient().Get(url)
	if err != nil {
		t.Fatalf("GET %s failed: %v", url, err)
	}
	defer resp.Body.Close()
	body, err := io.ReadAll(resp.Body)
	if err != nil {
		t.Fatalf("reading body from %s: %v", url, err)
	}
	return resp.StatusCode, string(body)
}

// assertContains fails the test when body does not contain the expected substring.
func assertContains(t *testing.T, body, expected string) {
	t.Helper()
	if !strings.Contains(body, expected) {
		t.Errorf("expected body to contain %q\nbody:\n%s", expected, body)
	}
}

// assertNotContains fails when body unexpectedly contains s.
func assertNotContains(t *testing.T, body, unexpected string) {
	t.Helper()
	if strings.Contains(body, unexpected) {
		t.Errorf("expected body NOT to contain %q\nbody:\n%s", unexpected, body)
	}
}

// skipIfNoServer skips the test when TEST_SERVER_URL is not set.
// Docker integration tests are opt-in to avoid failures for developers
// without the server running. CI sets TEST_SERVER_URL to enable them.
func skipIfNoServer(t *testing.T) {
	if os.Getenv("TEST_SERVER_URL") == "" {
		t.Skip("Skipping: TEST_SERVER_URL not set (Docker integration tests are opt-in)")
	}
}

// ─── Connectivity ─────────────────────────────────────────────────────────────

// TestDockerServerResponds verifies the server starts and the root URL is reachable.
func TestDockerServerResponds(t *testing.T) {
	skipIfNoServer(t)
	status, body := get(t, "/")
	if status != http.StatusOK {
		t.Errorf("expected HTTP 200 from /, got %d\nbody: %s", status, body)
	}
	if len(body) == 0 {
		t.Error("expected non-empty response body from /")
	}
}

// ─── Static file serving ──────────────────────────────────────────────────────

// TestDockerStaticSVG verifies that static files (non-ASP) are served correctly.
func TestDockerStaticSVG(t *testing.T) {
	skipIfNoServer(t)
	status, body := get(t, "/axonasp.svg")
	if status != http.StatusOK {
		t.Errorf("expected HTTP 200 for /axonasp.svg, got %d", status)
	}
	assertContains(t, body, "<svg")
}

// ─── Error pages ──────────────────────────────────────────────────────────────

// TestDockerNotFoundReturns404 verifies the 404 error page is served for missing paths.
func TestDockerNotFoundReturns404(t *testing.T) {
	skipIfNoServer(t)
	status, _ := get(t, "/this-path-does-not-exist-xyz.asp")
	if status != http.StatusNotFound {
		t.Errorf("expected HTTP 404 for missing resource, got %d", status)
	}
}

// TestDockerBlockedExtensionReturns404 verifies blocked file types return 404.
func TestDockerBlockedExtensionReturns404(t *testing.T) {
	skipIfNoServer(t)
	for _, ext := range []string{".env", ".config", ".cs", ".dll"} {
		path := fmt.Sprintf("/blocked-test%s", ext)
		status, _ := get(t, path)
		if status != http.StatusNotFound {
			t.Errorf("expected HTTP 404 for blocked extension %s, got %d", ext, status)
		}
	}
}

// ─── ASP execution ────────────────────────────────────────────────────────────

// TestDockerASPHelloWorld verifies basic ASP execution (Response.Write) works.
func TestDockerASPHelloWorld(t *testing.T) {
	skipIfNoServer(t)
	status, body := get(t, "/tests/test_hello.asp")
	if status != http.StatusOK {
		t.Errorf("expected HTTP 200, got %d", status)
	}
	assertContains(t, body, "Hello World")
}

// TestDockerASPSimple verifies that the simplest possible ASP page renders.
func TestDockerASPSimple(t *testing.T) {
	skipIfNoServer(t)
	status, body := get(t, "/tests/test_simple.asp")
	if status != http.StatusOK {
		t.Errorf("expected HTTP 200, got %d", status)
	}
	if len(body) == 0 {
		t.Error("expected non-empty body from test_simple.asp")
	}
}

// TestDockerASPEnvironment runs the dedicated Docker validation ASP page and
// confirms all VBScript feature checks report [PASS].
func TestDockerASPEnvironment(t *testing.T) {
	skipIfNoServer(t)
	status, body := get(t, "/tests/test_docker.asp")
	if status != http.StatusOK {
		t.Fatalf("expected HTTP 200 from test_docker.asp, got %d\nbody: %s", status, body)
	}

	// Verify no individual test failed
	lines := strings.Split(body, "\n")
	for _, line := range lines {
		line = strings.TrimSpace(line)
		if strings.HasPrefix(line, "[FAIL]") {
			t.Errorf("ASP test failed: %s", line)
		}
	}

	// Verify the summary line
	assertContains(t, body, "RESULT: ALL_PASS")
	assertNotContains(t, body, "RESULT: SOME_FAIL")
}

// TestDockerASPEnvironmentChecks validates individual [PASS] tokens from test_docker.asp.
func TestDockerASPEnvironmentChecks(t *testing.T) {
	skipIfNoServer(t)
	status, body := get(t, "/tests/test_docker.asp")
	if status != http.StatusOK {
		t.Fatalf("expected HTTP 200, got %d", status)
	}

	expectedChecks := []string{
		"[PASS] VBScript arithmetic",
		"[PASS] String concatenation",
		"[PASS] UCase",
		"[PASS] LCase",
		"[PASS] Len",
		"[PASS] Mid",
		"[PASS] InStr",
		"[PASS] Replace",
		"[PASS] Integer division",
		"[PASS] Mod",
		"[PASS] Abs",
		"[PASS] For loop",
		"[PASS] Do While loop",
		"[PASS] Select Case",
		"[PASS] User function",
		"[PASS] Sub call",
		"[PASS] Server.URLEncode",
		"[PASS] Server.MapPath",
		"[PASS] Server.HTMLEncode",
		"[PASS] Array index",
		"[PASS] UBound",
		"[PASS] ReDim Preserve",
		"[PASS] Class instantiation",
		"[PASS] Class method",
		"[PASS] On Error Resume Next",
		"[PASS] Dictionary Add/Item",
		"[PASS] Dictionary Count",
		"[PASS] Dictionary Exists",
	}

	for _, check := range expectedChecks {
		if !strings.Contains(body, check) {
			t.Errorf("missing expected check in test_docker.asp output: %s", check)
		}
	}
}

// ─── HTTP fundamentals ────────────────────────────────────────────────────────

// TestDockerResponseHeaders verifies that ASP pages set a Content-Type header.
func TestDockerResponseHeaders(t *testing.T) {
	skipIfNoServer(t)
	resp, err := newClient().Get(baseURL() + "/tests/test_hello.asp")
	if err != nil {
		t.Fatalf("request failed: %v", err)
	}
	defer resp.Body.Close()

	ct := resp.Header.Get("Content-Type")
	if ct == "" {
		t.Error("expected Content-Type header to be set")
	}
	if !strings.Contains(strings.ToLower(ct), "html") && !strings.Contains(strings.ToLower(ct), "text") {
		t.Errorf("unexpected Content-Type: %s", ct)
	}
}

// TestDockerDirectoryTraversalBlocked verifies that path traversal is rejected.
func TestDockerDirectoryTraversalBlocked(t *testing.T) {
	skipIfNoServer(t)
	paths := []string{
		"/../etc/passwd",
		"/tests/../../main.go",
		"/%2e%2e/go.mod",
	}
	for _, path := range paths {
		status, _ := get(t, path)
		if status == http.StatusOK {
			t.Errorf("path traversal attempt %q returned 200, expected 4xx", path)
		}
	}
}

// TestDockerConcurrentRequests verifies the server handles multiple simultaneous requests.
func TestDockerConcurrentRequests(t *testing.T) {
	skipIfNoServer(t)
	const workers = 10
	results := make(chan error, workers)

	for i := 0; i < workers; i++ {
		go func() {
			status, _ := get(t, "/tests/test_hello.asp")
			if status != http.StatusOK {
				results <- fmt.Errorf("expected 200, got %d", status)
			} else {
				results <- nil
			}
		}()
	}

	for i := 0; i < workers; i++ {
		if err := <-results; err != nil {
			t.Errorf("concurrent request failed: %v", err)
		}
	}
}

// TestDockerHealthEndpoint verifies the root path always returns a success response
// (serves as a basic health-check target).
func TestDockerHealthEndpoint(t *testing.T) {
	skipIfNoServer(t)
	for attempt := 1; attempt <= 3; attempt++ {
		status, _ := get(t, "/")
		if status == http.StatusOK {
			return
		}
		t.Logf("attempt %d: root returned %d, retrying...", attempt, status)
		time.Sleep(time.Duration(attempt) * 500 * time.Millisecond)
	}
	t.Error("server root did not return HTTP 200 after 3 attempts")
}
