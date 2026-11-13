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
package asp

import (
	"bytes"
	"net/http"
	"net/http/httptest"
	"strings"
	"testing"
)

// TestResponseSetCookieSubKeyPreservesIncomingRequestSubKeys verifies that modifying or adding
// one subkey on an existing request dictionary cookie preserves all other incoming subkeys
// in deterministic order and serializes them into Set-Cookie.
func TestResponseSetCookieSubKeyPreservesIncomingRequestSubKeys(t *testing.T) {
	req := httptest.NewRequest(http.MethodGet, "http://example.local/", nil)
	req.Header.Set("Cookie", "MGSsession=key1=val1&key2=val2")

	rec := httptest.NewRecorder()
	response := NewResponse(rec)
	defer response.ReleaseBuffer()
	response.SetRequest(req)

	response.SetCookieSubKey("MGSsession", "key3", "val3")
	if got := response.GetCookieValue("MGSsession"); got != "key1=val1&key2=val2&key3=val3" {
		t.Fatalf("unexpected merged cookie value: got %q want %q", got, "key1=val1&key2=val2&key3=val3")
	}

	// Repeated assignment in the same request updates key1 in place without losing key2/key3.
	response.SetCookieSubKey("MGSsession", "key1", "updated1")
	if got := response.GetCookieValue("MGSsession"); got != "key1=updated1&key2=val2&key3=val3" {
		t.Fatalf("unexpected updated cookie value: got %q want %q", got, "key1=updated1&key2=val2&key3=val3")
	}
	if !response.HasCookieKeys("MGSsession") {
		t.Fatalf("expected HasCookieKeys to return true")
	}
	if got := response.GetCookieProperty("MGSsession", "HasKeys"); got != "True" {
		t.Fatalf("expected HasKeys property to be True, got %q", got)
	}
	if got := response.GetCookieSubKey("MGSsession", "KEY2"); got != "val2" {
		t.Fatalf("expected KEY2=val2, got %q", got)
	}

	keys := response.GetCookieSubKeys("MGSsession")
	if strings.Join(keys, ",") != "key1,key2,key3" {
		t.Fatalf("unexpected subkey order: %v", keys)
	}

	response.Flush()
	setCookies := rec.Header().Values("Set-Cookie")
	if len(setCookies) != 1 {
		t.Fatalf("expected 1 Set-Cookie header, got %d (%v)", len(setCookies), setCookies)
	}
	if !strings.Contains(setCookies[0], "MGSsession=key1=updated1&key2=val2&key3=val3") {
		t.Fatalf("expected persisted subkeys in Set-Cookie header, got %q", setCookies[0])
	}
}

// TestResponseSetCookieSubKeyCaseInsensitiveIncomingLookup verifies that incoming
// HTTP cookie names are matched case-insensitively when loading existing subkeys.
func TestResponseSetCookieSubKeyCaseInsensitiveIncomingLookup(t *testing.T) {
	req := httptest.NewRequest(http.MethodGet, "http://example.local/", nil)
	req.Header.Set("Cookie", "MGSsession=uid=42&role=admin")

	var output bytes.Buffer
	response := NewResponse(&output)
	defer response.ReleaseBuffer()
	response.SetRequest(req)

	response.SetCookieSubKey("mgssession", "theme", "dark")
	if got := response.GetCookieValue("mgssession"); got != "uid=42&role=admin&theme=dark" {
		t.Fatalf("expected case-insensitive incoming subkey load, got %q", got)
	}
}

// TestResponseSetCookiePropertyBeforeSubKeyPreservesIncomingSubKeys verifies that
// setting cookie attributes such as Path or HttpOnly before assigning a subkey still
// loads existing incoming request subkeys.
func TestResponseSetCookiePropertyBeforeSubKeyPreservesIncomingSubKeys(t *testing.T) {
	req := httptest.NewRequest(http.MethodGet, "http://example.local/", nil)
	req.Header.Set("Cookie", "MGSsession=key1=val1")

	var output bytes.Buffer
	response := NewResponse(&output)
	defer response.ReleaseBuffer()
	response.SetRequest(req)

	response.SetCookieProperty("MGSsession", "Path", "/app")
	response.SetCookieProperty("MGSsession", "HttpOnly", "true")
	response.SetCookieSubKey("MGSsession", "key2", "val2")

	if got := response.GetCookieValue("MGSsession"); got != "key1=val1&key2=val2" {
		t.Fatalf("expected incoming subkeys to persist after property assignment, got %q", got)
	}
	if got := response.GetCookieProperty("MGSsession", "Path"); got != "/app" {
		t.Fatalf("expected Path=/app, got %q", got)
	}
	if got := response.GetCookieProperty("MGSsession", "HttpOnly"); got != "True" {
		t.Fatalf("expected HttpOnly=True, got %q", got)
	}
}

// TestResponseExplicitClearBeforeSubKeyDoesNotRestoreIncomingSubKeys verifies that
// explicitly clearing a response cookie before writing subkeys prevents incoming
// request subkeys from leaking back in.
func TestResponseExplicitClearBeforeSubKeyDoesNotRestoreIncomingSubKeys(t *testing.T) {
	req := httptest.NewRequest(http.MethodGet, "http://example.local/", nil)
	req.Header.Set("Cookie", "MGSsession=key1=val1&key2=val2")

	var output bytes.Buffer
	response := NewResponse(&output)
	defer response.ReleaseBuffer()
	response.SetRequest(req)

	response.SetCookieValue("MGSsession", "")
	response.SetCookieSubKey("MGSsession", "key3", "val3")

	if got := response.GetCookieValue("MGSsession"); got != "key3=val3" {
		t.Fatalf("expected cleared cookie to start fresh with key3=val3, got %q", got)
	}
}
