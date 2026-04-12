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
	"path/filepath"
	"runtime"
	"strings"
	"testing"
)

// TestResolveServiceExecutablePathAppendsExtensionOnWindows verifies automatic .exe suffix behavior.
func TestResolveServiceExecutablePathAppendsExtensionOnWindows(t *testing.T) {
	if runtime.GOOS != "windows" {
		t.Skip("windows-specific behavior")
	}

	result, err := resolveServiceExecutablePath(`C:\Axon`, "./axonasp-http")
	if err != nil {
		t.Fatalf("resolveServiceExecutablePath returned error: %v", err)
	}
	if filepath.Ext(result) != ".exe" {
		t.Fatalf("expected .exe extension, got %q", result)
	}
}

// TestResolveServiceExecutablePathPreservesExtension verifies explicit extension values are preserved.
func TestResolveServiceExecutablePathPreservesExtension(t *testing.T) {
	result, err := resolveServiceExecutablePath("/opt/axonasp", "./axonasp-fastcgi.bin")
	if err != nil {
		t.Fatalf("resolveServiceExecutablePath returned error: %v", err)
	}
	if !strings.HasSuffix(result, "axonasp-fastcgi.bin") {
		t.Fatalf("expected preserved extension, got %q", result)
	}
}

// TestMergeServiceEnvironmentRejectsInvalidEntry verifies KEY=VALUE validation.
func TestMergeServiceEnvironmentRejectsInvalidEntry(t *testing.T) {
	_, err := mergeServiceEnvironment([]string{"BROKEN"})
	if err == nil {
		t.Fatal("expected validation error for malformed entry")
	}
}

// TestMergeServiceEnvironmentAllowsValidEntries verifies valid entries are appended.
func TestMergeServiceEnvironmentAllowsValidEntries(t *testing.T) {
	merged, err := mergeServiceEnvironment([]string{"A=1", "B=2"})
	if err != nil {
		t.Fatalf("mergeServiceEnvironment returned error: %v", err)
	}

	if len(merged) < 2 {
		t.Fatalf("expected merged environment to include configured entries, got %d", len(merged))
	}
	if merged[len(merged)-2] != "A=1" || merged[len(merged)-1] != "B=2" {
		t.Fatalf("expected appended configured entries, got tail: %v", merged[len(merged)-2:])
	}
}
