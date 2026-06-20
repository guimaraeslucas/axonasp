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
	"path/filepath"
	"reflect"
	"testing"
)

// TestPerformanceFixItem1_ShareImmutableState verifies that VM initialization
// and re-initialization (resetForReuse) use the same backing array for global
// names and constants as the source program, avoiding heap allocations.
func TestPerformanceFixItem1_ShareImmutableState(t *testing.T) {
	compiler := NewASPCompiler(`<% Dim a, b: a = 1: b = 2 %>`)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	program := cachedProgramFromCompiler(compiler)
	vm := NewVMFromCachedProgram(program)

	// 1. Verify VM shared globalNames with program during initialization
	if getSliceBackingPtr(vm.globalNames) != getSliceBackingPtr(program.GlobalNames) {
		t.Errorf("expected vm.globalNames to share backing array with program.GlobalNames")
	}

	// 2. Verify VM shared globalNamesLower
	if getSliceBackingPtr(vm.baseGlobalNamesLower) != getSliceBackingPtr(program.GlobalNamesLower) {
		t.Errorf("expected vm.baseGlobalNamesLower to share backing array with program.GlobalNamesLower")
	}

	// 3. Verify captureBaseProgramState shared baseGlobalNames
	if getSliceBackingPtr(vm.baseGlobalNames) != getSliceBackingPtr(vm.globalNames) {
		t.Errorf("expected vm.baseGlobalNames to share backing array with vm.globalNames")
	}

	// 4. Verify resetForReuse restores sharing
	// First, simulate a mutation (e.g. from ExecuteGlobal)
	originalGlobalNames := vm.globalNames
	vm.globalNames = append(vm.globalNames, "new_global") // This will allocate a new backing array
	if getSliceBackingPtr(vm.globalNames) == getSliceBackingPtr(originalGlobalNames) {
		t.Errorf("mutation should have allocated a new backing array (cap was len)")
	}

	vm.resetForReuse()
	if getSliceBackingPtr(vm.globalNames) != getSliceBackingPtr(vm.baseGlobalNames) {
		t.Errorf("expected vm.globalNames to share backing array with vm.baseGlobalNames after reset")
	}
}

// getSliceBackingPtr returns the pointer to the underlying array of a slice.
func getSliceBackingPtr(slice any) uintptr {
	return reflect.ValueOf(slice).Pointer()
}

// TestPerformanceFixItem1_PoolLimit verifies the pool limit is set to 250.
func TestPerformanceFixItem1_PoolLimit(t *testing.T) {
	if vmProgramPoolDefaultRetained != 250 {
		t.Errorf("expected vmProgramPoolDefaultRetained to be 250, got %d", vmProgramPoolDefaultRetained)
	}
}

// TestPerformanceFixItem3_FSOResolvePathCache verifies that path resolution is memoized.
func TestPerformanceFixItem3_FSOResolvePathCache(t *testing.T) {
	// We need a VM with a mock host that supports MapPath
	vm := NewVM(nil, nil, 0)
	host := NewMockHost()
	vm.SetHost(host)

	path := "test.asp"

	// First call should populate cache
	res1, ok1 := vm.fsoResolvePath(path)
	if !ok1 {
		t.Fatalf("fsoResolvePath failed")
	}

	// Check if it's in pathCache
	rootPath := vm.host.Server().MapPath("/")
	currentDir := filepath.Dir(vm.host.Server().GetRequestPath())
	cacheKey := rootPath + "|" + currentDir + "|" + path

	if _, ok := globalFSOCache.pathCache.Load(cacheKey); !ok {
		t.Errorf("expected path to be cached in globalFSOCache.pathCache")
	}

	// Second call should come from cache (we can verify by checking if it returns the same string instance)
	// But in Go, strings are immutable, so we'll just verify it still works.
	res2, ok2 := vm.fsoResolvePath(path)
	if !ok2 || res2 != res1 {
		t.Errorf("fsoResolvePath from cache failed or returned different result")
	}
}
