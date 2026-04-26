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
package axonvm

import (
	"bytes"
	"strings"
	"testing"
)

// TestASPDirectiveCodePage verifies that page directives update Response.CodePage.
func TestASPDirectiveCodePage(t *testing.T) {
	source := `<%@ Language="VBScript" CODEPAGE="65001" %><%= Response.CodePage %>`
	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	host.Response().SetBuffer(false)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}

	if output.String() != "65001" {
		t.Fatalf("unexpected directive output: %q", output.String())
	}
	if host.Response().GetCodePage() != 65001 {
		t.Fatalf("unexpected response code page: %d", host.Response().GetCodePage())
	}
}

// TestASPDirectiveLanguageJScript verifies that Language="JScript" is accepted and executes server-side JScript.
func TestASPDirectiveLanguageJScript(t *testing.T) {
	source := `<%@ Language="JScript" %><%= "ok" %>`
	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}

	if output.String() != "ok" {
		t.Fatalf("unexpected JScript directive output: %q", output.String())
	}
}

// TestASPDirectiveDisableSession verifies that EnableSessionState=False disables Session access.
func TestASPDirectiveDisableSession(t *testing.T) {
	source := `<%@ EnableSessionState="False" %><%= "ok" %>`
	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}

	if host.SessionEnabled() {
		t.Fatal("expected session state to be disabled by directive")
	}

	source = `<%@ EnableSessionState="False" %><%= Session.Timeout %>`
	compiler = NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm = NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host = NewMockHost()
	vm.SetHost(host)

	err := vm.Run()
	if err == nil {
		t.Fatal("expected runtime error when Session is used while disabled")
	}
	if !strings.Contains(err.Error(), "Session state is disabled") {
		t.Fatalf("unexpected runtime error: %v", err)
	}
}
