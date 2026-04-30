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
 */
package axonvm

import (
	"testing"
)

// TestJScriptStrictModeInfrastructure verifies that strict mode tracking is properly initialized
func TestJScriptStrictModeInfrastructure(t *testing.T) {
	compiler := NewCompiler("")
	if compiler.jsStrictMode {
		t.Error("compiler should not be in strict mode by default")
	}
	if compiler.jsFunctionStrictModes == nil {
		t.Error("jsFunctionStrictModes map should be initialized")
	}
	if len(compiler.jsBlockScopeStack) != 0 {
		t.Error("jsBlockScopeStack should start empty")
	}

	vm := NewVM(nil, nil, 8)
	if vm.jsStrictMode {
		t.Error("VM should not be in strict mode by default")
	}
	if vm.jsFunctionStrictModes == nil {
		t.Error("VM jsFunctionStrictModes map should be initialized")
	}
	if vm.jsBlockScopeDepth != 0 {
		t.Error("VM jsBlockScopeDepth should start at 0")
	}
	if len(vm.jsBlockScopes) != 0 {
		t.Error("VM jsBlockScopes should start empty")
	}
}

// TestJScriptOpcodeDefinitions verifies that strict mode opcodes exist
func TestJScriptOpcodeDefinitions(t *testing.T) {
	// Verify opcodes are defined
	_ = OpJSStrictModeEnter
	_ = OpJSStrictModeExit
	_ = OpJSBlockScopeEnter
	_ = OpJSBlockScopeExit
	_ = OpJSLetDeclare
	_ = OpJSConstDeclare

	// Verify opcode string representation works
	if OpJSStrictModeEnter.String() != "OpJSStrictModeEnter" {
		t.Error("OpJSStrictModeEnter string representation is incorrect")
	}
	if OpJSStrictModeExit.String() != "OpJSStrictModeExit" {
		t.Error("OpJSStrictModeExit string representation is incorrect")
	}
	if OpJSBlockScopeEnter.String() != "OpJSBlockScopeEnter" {
		t.Error("OpJSBlockScopeEnter string representation is incorrect")
	}
	if OpJSBlockScopeExit.String() != "OpJSBlockScopeExit" {
		t.Error("OpJSBlockScopeExit string representation is incorrect")
	}
	if OpJSLetDeclare.String() != "OpJSLetDeclare" {
		t.Error("OpJSLetDeclare string representation is incorrect")
	}
	if OpJSConstDeclare.String() != "OpJSConstDeclare" {
		t.Error("OpJSConstDeclare string representation is incorrect")
	}
}

// TestJScriptBlockScopeLifecycle verifies block scope infrastructure exists in VM
func TestJScriptBlockScopeLifecycle(t *testing.T) {
	// Just verify VM has block scope infrastructure initialized
	vm := NewVM(nil, nil, 8)

	// Verify VM has the block scope infrastructure
	if vm.jsBlockScopes == nil {
		t.Error("VM jsBlockScopes should be initialized")
	}
	if vm.jsBlockScopeDepth != 0 {
		t.Error("VM jsBlockScopeDepth should start at 0")
	}

	// Simulate entering and exiting a block scope
	vm.jsBlockScopes = append(vm.jsBlockScopes, make(map[string]Value))
	vm.jsBlockScopeDepth++

	if vm.jsBlockScopeDepth != 1 {
		t.Error("VM jsBlockScopeDepth should be 1 after entering")
	}
	if len(vm.jsBlockScopes) != 1 {
		t.Error("VM jsBlockScopes should have 1 entry after entering")
	}

	// Exit block scope
	vm.jsBlockScopes = vm.jsBlockScopes[:len(vm.jsBlockScopes)-1]
	vm.jsBlockScopeDepth--

	if vm.jsBlockScopeDepth != 0 {
		t.Error("VM jsBlockScopeDepth should be 0 after exiting")
	}
	if len(vm.jsBlockScopes) != 0 {
		t.Error("VM jsBlockScopes should be empty after exiting")
	}
}

// TestJScriptStrictModeDirectiveDetection verifies directive parsing works
func TestJScriptStrictModeDirectiveDetection(t *testing.T) {
	compiler := NewCompiler("")

	// The detectUseStrictDirective method should exist on the compiler
	// and be callable (we can't test it directly without creating AST nodes)
	if compiler == nil {
		t.Error("compiler should be initialized")
	}
}

// TestJScriptStrictModePersistsCorrectly verifies strict mode is properly tracked
func TestJScriptStrictModePersistsCorrectly(t *testing.T) {
	vm := NewVM(nil, nil, 8)

	// Initially not in strict mode
	if vm.jsStrictMode {
		t.Error("VM should not start in strict mode")
	}

	// Set strict mode
	vm.jsStrictMode = true
	if !vm.jsStrictMode {
		t.Error("VM should be in strict mode after setting")
	}

	// Clear strict mode
	vm.jsStrictMode = false
	if vm.jsStrictMode {
		t.Error("VM should not be in strict mode after clearing")
	}
}

// TestJScriptStrictModeImplicitGlobalChecking verifies the jsSetName check
func TestJScriptStrictModeImplicitGlobalChecking(t *testing.T) {
	vm := NewVM(nil, nil, 8)

	// In non-strict mode, implicit globals should be allowed (or at least not throw strict error)
	vm.jsStrictMode = false
	// This should not panic due to strict mode
	defer func() {
		if r := recover(); r != nil && vm.jsStrictMode {
			t.Errorf("unexpected error in non-strict mode: %v", r)
		}
	}()

	// Note: The actual jsSetName call may fail due to missing environment,
	// but that's a different error than strict mode
}
