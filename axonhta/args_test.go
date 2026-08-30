/*
 * AxonASP Server
 * Copyright (C) 2026 G3pix Ltda. All rights reserved.
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 */
package main

import (
	"os"
	"path/filepath"
	"testing"
)

// TestResolveHTAArg_NoArgs verifies that resolveHTAArg is a no-op when
// called with only the program name in os.Args.
func TestResolveHTAArg_NoArgs(t *testing.T) {
	orig := os.Args
	defer func() { os.Args = orig }()

	htaFileOverride = ""
	appDir = "./"
	os.Args = []string{"axonhta.exe"}

	resolveHTAArg()

	if htaFileOverride != "" {
		t.Errorf("expected htaFileOverride to be empty, got %q", htaFileOverride)
	}
}

// TestResolveHTAArg_ValidHTAFile verifies that a valid existing .hta file
// passed as a positional argument sets htaFileOverride and appDir.
func TestResolveHTAArg_ValidHTAFile(t *testing.T) {
	// Create a temporary .hta file.
	tmp, err := os.CreateTemp(t.TempDir(), "*.hta")
	if err != nil {
		t.Fatal(err)
	}
	tmp.Close()

	orig := os.Args
	defer func() {
		os.Args = orig
		htaFileOverride = ""
		appDir = "./"
	}()

	htaFileOverride = ""
	appDir = "./"
	os.Args = []string{"axonhta.exe", tmp.Name()}

	resolveHTAArg()

	if htaFileOverride == "" {
		t.Fatal("expected htaFileOverride to be set")
	}

	wantDir := filepath.Dir(tmp.Name())
	if appDir != wantDir {
		t.Errorf("appDir: want %q, got %q", wantDir, appDir)
	}
}

// TestResolveHTAArg_NonHTAExtension verifies that non-.hta files are ignored.
func TestResolveHTAArg_NonHTAExtension(t *testing.T) {
	tmp, err := os.CreateTemp(t.TempDir(), "*.txt")
	if err != nil {
		t.Fatal(err)
	}
	tmp.Close()

	orig := os.Args
	defer func() {
		os.Args = orig
		htaFileOverride = ""
		appDir = "./"
	}()

	htaFileOverride = ""
	appDir = "./"
	os.Args = []string{"axonhta.exe", tmp.Name()}

	resolveHTAArg()

	if htaFileOverride != "" {
		t.Errorf("expected htaFileOverride to remain empty for non-.hta file, got %q", htaFileOverride)
	}
}

// TestResolveHTAArg_MissingFile verifies that a non-existent .hta path is
// silently ignored and does not set htaFileOverride.
func TestResolveHTAArg_MissingFile(t *testing.T) {
	orig := os.Args
	defer func() {
		os.Args = orig
		htaFileOverride = ""
		appDir = "./"
	}()

	htaFileOverride = ""
	appDir = "./"
	os.Args = []string{"axonhta.exe", "/nonexistent/path/to/file.hta"}

	resolveHTAArg()

	if htaFileOverride != "" {
		t.Errorf("expected htaFileOverride to remain empty for missing file, got %q", htaFileOverride)
	}
}

// TestResolveHTAArg_FlagSkipped verifies that flag-style arguments (starting
// with "-") are not mistakenly interpreted as file paths.
func TestResolveHTAArg_FlagSkipped(t *testing.T) {
	orig := os.Args
	defer func() {
		os.Args = orig
		htaFileOverride = ""
		appDir = "./"
	}()

	htaFileOverride = ""
	appDir = "./"
	os.Args = []string{"axonhta.exe", "-dev", "--app=./myapp"}

	resolveHTAArg()

	if htaFileOverride != "" {
		t.Errorf("expected htaFileOverride to remain empty for flag args, got %q", htaFileOverride)
	}
}
