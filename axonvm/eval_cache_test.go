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

import "testing"

// buildEvalCacheTestVM creates one VM with initialized compiler/runtime metadata for Eval cache tests.
func buildEvalCacheTestVM(t *testing.T) *VM {
	t.Helper()

	compiler := NewASPCompiler(`<% Response.Write "" %>`)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}
	return NewVMFromCompiler(compiler)
}

// TestEvalCacheHitSameScope verifies repeated Eval compilation requests reuse one cached payload.
func TestEvalCacheHitSameScope(t *testing.T) {
	vm := buildEvalCacheTestVM(t)
	localSub := Value{Type: VTUserSub, Names: []string{"valueProbe"}}

	first, err := vm.getOrCompileEvalProgram("valueProbe", localSub)
	if err != nil {
		t.Fatalf("first compile failed: %v", err)
	}
	second, err := vm.getOrCompileEvalProgram("valueProbe", localSub)
	if err != nil {
		t.Fatalf("second compile failed: %v", err)
	}

	if first == nil || second == nil {
		t.Fatalf("expected cached eval programs to be non-nil")
	}
	if first != second {
		t.Fatalf("expected eval cache hit to reuse same compiled payload")
	}
}

// TestEvalCacheMissDifferentLocalScope verifies local-frame layout changes invalidate cached Eval bytecode.
func TestEvalCacheMissDifferentLocalScope(t *testing.T) {
	vm := buildEvalCacheTestVM(t)

	scopeA := Value{Type: VTUserSub, Names: []string{"valueProbe"}}
	scopeB := Value{Type: VTUserSub, Names: []string{"tempProbe", "valueProbe"}}

	first, err := vm.getOrCompileEvalProgram("valueProbe", scopeA)
	if err != nil {
		t.Fatalf("scopeA compile failed: %v", err)
	}
	second, err := vm.getOrCompileEvalProgram("valueProbe", scopeB)
	if err != nil {
		t.Fatalf("scopeB compile failed: %v", err)
	}

	if first == nil || second == nil {
		t.Fatalf("expected eval programs to be non-nil")
	}
	if first == second {
		t.Fatalf("expected different local scopes to bypass cached eval payload reuse")
	}
	if first.localScopeHash == second.localScopeHash {
		t.Fatalf("expected different local scope fingerprints")
	}
}
