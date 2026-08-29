//go:build !lib_g3axonlive_disabled && !wasm

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
	"net/http"
	"net/http/httptest"
	"testing"

	"g3pix.com.br/axonasp/v2/axonvm"
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

// TestResolveG3ALScriptURL verifies page lookup uses the authenticated session mapping.
func TestResolveG3ALScriptURL(t *testing.T) {
	authSession := "RESOLVE123"
	scriptURL := "/axonlive/counter.asp"
	axonvm.G3ALRegisterPage(authSession, scriptURL)

	resolved := resolveG3ALScriptURL(authSession)
	if resolved != scriptURL {
		t.Fatalf("expected script URL %q, got %q", scriptURL, resolved)
	}
}
