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
	"os"
	"testing"
)

func TestWScriptShell(t *testing.T) {
	vm := NewVM(nil, nil, 0)

	shellLib := vm.newWScriptShellObject()
	if shellLib.Type != VTNativeObject {
		t.Fatalf("expected VTNativeObject, got %v", shellLib.Type)
	}

	obj := vm.wscriptShellItems[shellLib.Num]
	if obj == nil {
		t.Fatal("expected object in vm items")
	}

	// Test GetEnv
	res := obj.DispatchMethod("GetEnv", []Value{NewString("PATH")})
	if res.Type != VTString {
		t.Fatalf("expected VTString from GetEnv, got %v", res.Type)
	}

	// Test Run
	// Note: We use a safe command like 'echo' or 'cmd /c exit 0' to avoid hanging
	// We'll skip actual Exec execution in tests to prevent cross-platform hang issues
}

// TestWScriptShellExpandEnvironmentStrings verifies %NAME% expansion compatibility.
func TestWScriptShellExpandEnvironmentStrings(t *testing.T) {
	const envName = "AXONASP_WSCRIPT_EXPAND_TEST"
	const envValue = "expand-ok"

	if err := os.Setenv(envName, envValue); err != nil {
		t.Fatalf("setenv failed: %v", err)
	}
	t.Cleanup(func() {
		_ = os.Unsetenv(envName)
	})

	vm := NewVM(nil, nil, 0)
	shellLib := vm.newWScriptShellObject()
	obj := vm.wscriptShellItems[shellLib.Num]
	if obj == nil {
		t.Fatal("expected object in vm items")
	}

	got := obj.DispatchMethod("ExpandEnvironmentStrings", []Value{NewString("prefix-%" + envName + "%-suffix")})
	if got.Type != VTString {
		t.Fatalf("expected VTString from ExpandEnvironmentStrings, got %v", got.Type)
	}
	if got.Str != "prefix-"+envValue+"-suffix" {
		t.Fatalf("unexpected expansion result: got %q", got.Str)
	}
}

// TestWScriptShellExpandEnvironmentStringsUnknownPreserved verifies unknown placeholders stay unchanged.
func TestWScriptShellExpandEnvironmentStringsUnknownPreserved(t *testing.T) {
	vm := NewVM(nil, nil, 0)
	shellLib := vm.newWScriptShellObject()
	obj := vm.wscriptShellItems[shellLib.Num]
	if obj == nil {
		t.Fatal("expected object in vm items")
	}

	got := obj.DispatchMethod("ExpandEnvironmentStrings", []Value{NewString("%__AXONASP_UNKNOWN_ENV__%")})
	if got.Type != VTString {
		t.Fatalf("expected VTString from ExpandEnvironmentStrings, got %v", got.Type)
	}
	if got.Str != "%__AXONASP_UNKNOWN_ENV__%" {
		t.Fatalf("unexpected unknown-placeholder handling: got %q", got.Str)
	}
}
