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
	"bytes"
	"strings"
	"testing"
)

// TestResolveConsoleOutputTarget_CLITUI verifies that TUI mode routes console output
// to the host output buffer and suppresses trailing newlines.
func TestResolveConsoleOutputTarget_CLITUI(t *testing.T) {
	host := NewMockHost()
	var out bytes.Buffer
	host.SetOutput(&out)
	host.Request().ServerVars.Add("AXONASP_CLI_TUI", "1")

	vm := NewVM(nil, nil, 0)
	vm.SetHost(host)

	writer, lineEnding := resolveConsoleOutputTarget(vm, consoleMethodFormats["log"])
	if lineEnding != "" {
		t.Fatalf("expected no line ending in TUI mode, got %q", lineEnding)
	}
	if writer != &out {
		t.Fatalf("expected TUI writer to be host output buffer")
	}

	consoleDispatch(vm, "log", []Value{NewString("hello")})
	rendered := out.String()
	if rendered == "" {
		t.Fatalf("expected console output in host buffer")
	}
	if strings.HasSuffix(rendered, "\n") {
		t.Fatalf("expected no trailing newline in TUI mode output, got %q", rendered)
	}
}

// TestResolveConsoleOutputTarget_Default verifies non-TUI behavior keeps newline output.
func TestResolveConsoleOutputTarget_Default(t *testing.T) {
	host := NewMockHost()
	vm := NewVM(nil, nil, 0)
	vm.SetHost(host)

	_, lineEnding := resolveConsoleOutputTarget(vm, consoleMethodFormats["log"])
	if lineEnding != "\n" {
		t.Fatalf("expected default line ending to be newline, got %q", lineEnding)
	}
}
