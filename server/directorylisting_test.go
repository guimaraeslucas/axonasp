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
	"encoding/base64"
	"net/http"
	"net/http/httptest"
	"os"
	"path/filepath"
	"strings"
	"testing"
	"time"
)

func TestDirectoryListingRendererFiltersBlockedEntries(t *testing.T) {
	root := t.TempDir()
	templatePath := filepath.Join(root, "listing.html")
	templateBody := `<!doctype html><html><body>{{range .Entries}}<div>{{.Name}}|{{.MimeType}}</div>{{end}}</body></html>`
	if err := os.WriteFile(templatePath, []byte(templateBody), 0o644); err != nil {
		t.Fatalf("write template: %v", err)
	}

	dirPath := filepath.Join(root, "public")
	if err := os.MkdirAll(filepath.Join(dirPath, "allowed-dir"), 0o755); err != nil {
		t.Fatalf("mkdir allowed-dir: %v", err)
	}
	if err := os.MkdirAll(filepath.Join(dirPath, "hidden-dir"), 0o755); err != nil {
		t.Fatalf("mkdir hidden-dir: %v", err)
	}
	if err := os.WriteFile(filepath.Join(dirPath, "readme.txt"), []byte("ok"), 0o644); err != nil {
		t.Fatalf("write readme.txt: %v", err)
	}
	if err := os.WriteFile(filepath.Join(dirPath, "secret.config"), []byte("hidden"), 0o644); err != nil {
		t.Fatalf("write secret.config: %v", err)
	}
	if err := os.WriteFile(filepath.Join(dirPath, "MyInfo.xml"), []byte("hidden"), 0o644); err != nil {
		t.Fatalf("write MyInfo.xml: %v", err)
	}

	origBlockedFiles := BlockedFiles
	origBlockedExt := BlockedExtensions
	origBlockedDirs := BlockedDirs
	origBlockedPrefixes := blockedDirPrefixes
	origTZ := serverLocation
	defer func() {
		BlockedFiles = origBlockedFiles
		BlockedExtensions = origBlockedExt
		BlockedDirs = origBlockedDirs
		blockedDirPrefixes = origBlockedPrefixes
		serverLocation = origTZ
	}()

	BlockedFiles = []string{"myinfo.xml"}
	BlockedExtensions = []string{".config"}
	BlockedDirs = []string{filepath.Join(dirPath, "hidden-dir")}
	blockedDirPrefixes = buildBlockedDirPrefixes(BlockedDirs)
	serverLocation = time.UTC

	renderer, err := NewDirectoryListingRenderer(root, templatePath)
	if err != nil {
		t.Fatalf("create renderer: %v", err)
	}

	recorder := httptest.NewRecorder()
	req := httptest.NewRequest(http.MethodGet, "http://example.local/public/", nil)
	if err := renderer.Render(recorder, req, dirPath, "/public/"); err != nil {
		t.Fatalf("render listing: %v", err)
	}

	body := recorder.Body.String()
	if strings.Contains(body, "hidden-dir") {
		t.Fatalf("expected hidden-dir to be filtered")
	}
	if strings.Contains(body, "secret.config") {
		t.Fatalf("expected secret.config to be filtered")
	}
	if strings.Contains(body, "MyInfo.xml") {
		t.Fatalf("expected MyInfo.xml to be filtered")
	}
	if !strings.Contains(body, "allowed-dir") {
		t.Fatalf("expected allowed-dir in listing")
	}
	if !strings.Contains(body, "readme.txt") {
		t.Fatalf("expected readme.txt in listing")
	}
	if !strings.Contains(strings.ToLower(body), "readme.txt|text/plain") {
		t.Fatalf("expected MIME type for readme.txt, got body: %s", body)
	}
}

func TestReadLogoAsDataURIOrEmpty(t *testing.T) {
	tmpDir := t.TempDir()
	logoPath := filepath.Join(tmpDir, "logo.png")
	logoData := []byte{0x89, 'P', 'N', 'G', '\r', '\n', 0x1A, '\n', 0x00, 0x00, 0x00, 0x00}
	if err := os.WriteFile(logoPath, logoData, 0o644); err != nil {
		t.Fatalf("write logo fixture: %v", err)
	}

	got := readLogoAsDataURIOrEmpty(logoPath)
	want := "data:image/png;base64," + base64.StdEncoding.EncodeToString(logoData)
	if got != want {
		t.Fatalf("unexpected data URI, want %q got %q", want, got)
	}
}

func TestReadTextFileOrEmpty(t *testing.T) {
	tmpDir := t.TempDir()
	cssPath := filepath.Join(tmpDir, "style.css")
	cssContent := "body { background: #ece9d8; }"
	if err := os.WriteFile(cssPath, []byte(cssContent), 0o644); err != nil {
		t.Fatalf("write css fixture: %v", err)
	}

	got := readTextFileOrEmpty(cssPath)
	if got != cssContent {
		t.Fatalf("unexpected css content, want %q got %q", cssContent, got)
	}
}
