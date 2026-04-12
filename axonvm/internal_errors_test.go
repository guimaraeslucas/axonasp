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
package axonvm

import (
	"errors"
	"strings"
	"testing"
)

// TestNewAxonASPErrorUsesCatalogMessage verifies that formatted internal errors include code and catalog text.
func TestNewAxonASPErrorUsesCatalogMessage(t *testing.T) {
	wrapped := errors.New("access denied")
	err := NewAxonASPError(ErrCouldNotResolveCurrentDir, wrapped, "", "cli/main.go", 42)

	message := err.Error()
	if !strings.Contains(message, "AxonASP Error [2003] Could not resolve current directory") {
		t.Fatalf("unexpected error message: %q", message)
	}
	if !strings.Contains(message, "File: cli/main.go") {
		t.Fatalf("expected file name in message: %q", message)
	}
	if !strings.Contains(message, "Line: 42") {
		t.Fatalf("expected line number in message: %q", message)
	}
	if !strings.Contains(message, "Cause: access denied") {
		t.Fatalf("expected wrapped cause in message: %q", message)
	}
}

// TestAsAxonASPError verifies that wrapped AxonASP errors can be extracted through the standard error chain.
func TestAsAxonASPError(t *testing.T) {
	err := NewAxonASPError(ErrFileNotFound, errors.New("missing"), "", "", 0)

	resolved, ok := AsAxonASPError(err)
	if !ok {
		t.Fatalf("expected AsAxonASPError to resolve the structured error")
	}
	if resolved.Code != ErrFileNotFound {
		t.Fatalf("unexpected error code: got %d want %d", resolved.Code, ErrFileNotFound)
	}
}
