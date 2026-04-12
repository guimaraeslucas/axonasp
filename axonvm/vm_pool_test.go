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

import "testing"

// TestAcquireVMFromCachedProgramResetsState verifies pooled VMs restore immutable program state
// and clear request-scoped data before being reused by another execution.
func TestAcquireVMFromCachedProgramResetsState(t *testing.T) {
	compiler := NewASPCompiler(`<% Response.Write "ok" %>`)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	program := cachedProgramFromCompiler(compiler)
	vm := AcquireVMFromCachedProgram(program)
	host := NewMockHost()
	vm.SetHost(host)
	vm.Globals[len(vm.Globals)-1] = NewString("dirty")
	vm.globalNames = append(vm.globalNames, "dynamicGlobal")
	vm.declaredGlobals["dynamicglobal"] = true
	vm.constGlobals["dynamicconst"] = true
	vm.responseCookieItems[20001] = "cookie"
	vm.nativeObjectProxies[20002] = nativeObjectProxy{ParentID: 1, Member: "Dirty"}
	vm.ip = 9
	vm.sp = 3
	vm.lastLine = 42
	vm.Release()

	reused := AcquireVMFromCachedProgram(program)
	defer reused.Release()

	if reused.host != nil {
		t.Fatalf("expected pooled VM host to be cleared")
	}
	if reused.ip != 0 || reused.sp != -1 || reused.fp != 0 {
		t.Fatalf("expected VM execution pointers to be reset, got ip=%d sp=%d fp=%d", reused.ip, reused.sp, reused.fp)
	}
	if reused.lastLine != 0 || reused.lastColumn != 0 || reused.lastError != nil {
		t.Fatalf("expected last error state to be reset")
	}
	expectedGlobalCount := len(getBaseGlobalDictionary().names) + len(program.GlobalPreludeNames) + len(program.UserGlobalNames)
	if len(reused.globalNames) != expectedGlobalCount {
		t.Fatalf("expected global names to be restored, got %d want %d", len(reused.globalNames), expectedGlobalCount)
	}
	if _, exists := reused.declaredGlobals["dynamicglobal"]; exists {
		t.Fatalf("expected declared globals to be reset")
	}
	if _, exists := reused.constGlobals["dynamicconst"]; exists {
		t.Fatalf("expected const globals to be reset")
	}
	if len(reused.responseCookieItems) != 0 || len(reused.nativeObjectProxies) != 0 {
		t.Fatalf("expected dynamic native-object maps to be cleared")
	}
	if reused.Globals[len(reused.Globals)-1].Type == VTString && reused.Globals[len(reused.Globals)-1].Str == "dirty" {
		t.Fatalf("expected globals to be restored from the base template")
	}
	if len(reused.Globals) >= 7 {
		if reused.Globals[0].Type != VTNativeObject || reused.Globals[0].Num != 0 {
			t.Fatalf("expected Response intrinsic to be restored")
		}
		if reused.Globals[4].Type != VTNativeObject || reused.Globals[4].Num != 4 {
			t.Fatalf("expected Application intrinsic to be restored")
		}
	}
}
