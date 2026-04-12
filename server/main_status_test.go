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
	"net/http"
	"net/http/httptest"
	"os"
	"path/filepath"
	"testing"
)

// TestHandleRequestServesStaticPNG verifies non-ASP image requests are served
// directly from the configured web root without entering ASP execution.
func TestHandleRequestServesStaticPNG(t *testing.T) {
	root := t.TempDir()
	imageDir := filepath.Join(root, "backsiteTemplate31", "images")
	if err := os.MkdirAll(imageDir, 0o755); err != nil {
		t.Fatalf("create image dir: %v", err)
	}

	pngBody := []byte{0x89, 'P', 'N', 'G', '\r', '\n', 0x1A, '\n'}
	imagePath := filepath.Join(imageDir, "menuitem.png")
	if err := os.WriteFile(imagePath, pngBody, 0o644); err != nil {
		t.Fatalf("write image: %v", err)
	}

	originalRootDir := RootDir
	originalDefaultPages := DefaultPages
	originalExecuteAsASP := ExecuteAsASPExtensions
	originalBlockedExtensions := BlockedExtensions
	originalBlockedFiles := BlockedFiles
	originalBlockedDirs := BlockedDirs
	originalBlockedPrefixes := blockedDirPrefixes
	originalWebConfig := activeWebConfig
	originalDirectoryListing := directoryListingRenderer
	defer func() {
		RootDir = originalRootDir
		DefaultPages = originalDefaultPages
		ExecuteAsASPExtensions = originalExecuteAsASP
		BlockedExtensions = originalBlockedExtensions
		BlockedFiles = originalBlockedFiles
		BlockedDirs = originalBlockedDirs
		blockedDirPrefixes = originalBlockedPrefixes
		activeWebConfig = originalWebConfig
		directoryListingRenderer = originalDirectoryListing
	}()

	RootDir = root
	DefaultPages = []string{"default.asp"}
	ExecuteAsASPExtensions = []string{".asp"}
	BlockedExtensions = nil
	BlockedFiles = nil
	BlockedDirs = nil
	blockedDirPrefixes = nil
	activeWebConfig = nil
	directoryListingRenderer = nil

	rec := httptest.NewRecorder()
	req := httptest.NewRequest(http.MethodGet, "http://example.local/backsiteTemplate31/images/menuitem.png", nil)

	handleRequest(rec, req)

	if rec.Code != http.StatusOK {
		t.Fatalf("expected status %d, got %d", http.StatusOK, rec.Code)
	}
	if got := rec.Header().Get("Content-Type"); got != "image/png" {
		t.Fatalf("expected image/png content type, got %q", got)
	}
	if body := rec.Body.Bytes(); string(body) != string(pngBody) {
		t.Fatalf("unexpected static body bytes: %v", body)
	}
}

// TestSingleHeaderResponseWriterOverrides200WithDefault verifies that configured
// default error statuses are not downgraded to 200 by wrapped handlers.
func TestSingleHeaderResponseWriterOverrides200WithDefault(t *testing.T) {
	rec := httptest.NewRecorder()
	w := newSingleHeaderResponseWriter(rec, http.StatusNotFound)

	w.WriteHeader(http.StatusOK)

	if rec.Code != http.StatusNotFound {
		t.Fatalf("expected status %d, got %d", http.StatusNotFound, rec.Code)
	}
}

// TestSingleHeaderResponseWriterPreservesNon200 verifies explicit non-200 statuses remain intact.
func TestSingleHeaderResponseWriterPreservesNon200(t *testing.T) {
	rec := httptest.NewRecorder()
	w := newSingleHeaderResponseWriter(rec, http.StatusNotFound)

	w.WriteHeader(http.StatusInternalServerError)

	if rec.Code != http.StatusInternalServerError {
		t.Fatalf("expected status %d, got %d", http.StatusInternalServerError, rec.Code)
	}
}

// TestServeErrorPageHTMLPreservesStatus verifies static HTML error handlers return
// the requested HTTP status code instead of implicit 200.
func TestServeErrorPageHTMLPreservesStatus(t *testing.T) {
	tmp := t.TempDir()
	errorDir := filepath.Join(tmp, "error-pages")
	if err := os.MkdirAll(errorDir, 0o755); err != nil {
		t.Fatalf("create error dir: %v", err)
	}
	htmlPath := filepath.Join(errorDir, "404.html")
	if err := os.WriteFile(htmlPath, []byte("not found"), 0o644); err != nil {
		t.Fatalf("write 404.html: %v", err)
	}

	originalDir := DefaultErrorPagesDirectory
	DefaultErrorPagesDirectory = errorDir
	defer func() {
		DefaultErrorPagesDirectory = originalDir
	}()

	rec := httptest.NewRecorder()
	req := httptest.NewRequest(http.MethodGet, "http://example.local/missing.asp", nil)

	serveErrorPage(rec, req, http.StatusNotFound)

	if rec.Code != http.StatusNotFound {
		t.Fatalf("expected status %d, got %d", http.StatusNotFound, rec.Code)
	}
	if body := rec.Body.String(); body == "" {
		t.Fatalf("expected non-empty error page body")
	}
}
