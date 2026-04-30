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
	"bytes"
	"strings"
	"testing"
)

// jscriptSrc wraps raw JScript source in the ASP server-side script tag.
func jscriptSrc(src string) string {
	return `<script runat="server" language="JScript">` + src + `</script>`
}

// runJScript2 compiles and runs a JScript ASP source string, returning
// (output, runError). Compile errors are returned as non-nil runError.
func runJScript2(t *testing.T, aspSrc string) (output string, runErr error) {
	t.Helper()
	compiler := NewASPCompiler(aspSrc)
	if err := compiler.Compile(); err != nil {
		return "", err
	}
	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var buf bytes.Buffer
	host.SetOutput(&buf)
	host.Response().SetBuffer(false)
	vm.SetHost(host)
	runErr = vm.Run()
	return buf.String(), runErr
}

// compileJScript2 compiles JScript ASP source and returns any compile error.
// Panics from the compiler (JSSyntaxError etc.) are also caught and returned.
func compileJScript2(aspSrc string) (compileErr error) {
	defer func() {
		if r := recover(); r != nil {
			if e, ok := r.(error); ok {
				compileErr = e
			} else {
				compileErr = &phase2TestError{r}
			}
		}
	}()
	compiler := NewASPCompiler(aspSrc)
	return compiler.Compile()
}

// phase2TestError wraps a non-error panic value recovered from the compiler.
type phase2TestError struct{ val interface{} }

func (e *phase2TestError) Error() string {
	return strings.TrimSpace(strings.ReplaceAll(strings.TrimSpace(func() string {
		if s, ok := e.val.(string); ok {
			return s
		}
		return "compiler panic"
	}()), "  ", " "))
}

// ─────────────────────────────────────────────────────────────────────────────
// Phase 2 opcode definitions
// ─────────────────────────────────────────────────────────────────────────────

// TestJScriptPhase2OpcodeDefinitions verifies that all three Phase 2 opcodes
// exist with correct String() representations.
func TestJScriptPhase2OpcodeDefinitions(t *testing.T) {
	tests := []struct {
		op   OpCode
		name string
	}{
		{OpJSConstInitialize, "OpJSConstInitialize"},
		{OpJSForIterEnter, "OpJSForIterEnter"},
		{OpJSForIterExit, "OpJSForIterExit"},
	}
	for _, tt := range tests {
		if got := tt.op.String(); got != tt.name {
			t.Errorf("opcode string: want %q, got %q", tt.name, got)
		}
	}
}

// TestJScriptPhase2VMFieldsInitialized verifies that jsBlockScopeConst and
// jsBlockScopeTDZ are properly initialized on a fresh VM.
func TestJScriptPhase2VMFieldsInitialized(t *testing.T) {
	vm := NewVM(nil, nil, 8)
	if vm.jsBlockScopeConst == nil {
		t.Error("jsBlockScopeConst should be initialized")
	}
	if vm.jsBlockScopeTDZ == nil {
		t.Error("jsBlockScopeTDZ should be initialized")
	}
	if len(vm.jsBlockScopeConst) != 0 {
		t.Error("jsBlockScopeConst should start empty")
	}
	if len(vm.jsBlockScopeTDZ) != 0 {
		t.Error("jsBlockScopeTDZ should start empty")
	}
}

// TestJScriptBlockScopeParallelSlices verifies that all three block scope slices
// stay in sync on simulated enter/exit operations.
func TestJScriptBlockScopeParallelSlices(t *testing.T) {
	vm := NewVM(nil, nil, 8)

	// Simulate OpJSBlockScopeEnter.
	vm.jsBlockScopes = append(vm.jsBlockScopes, make(map[string]Value, 4))
	vm.jsBlockScopeConst = append(vm.jsBlockScopeConst, make(map[string]struct{}, 2))
	vm.jsBlockScopeTDZ = append(vm.jsBlockScopeTDZ, make(map[string]struct{}, 2))
	vm.jsBlockScopeDepth++

	if len(vm.jsBlockScopes) != 1 || len(vm.jsBlockScopeConst) != 1 || len(vm.jsBlockScopeTDZ) != 1 {
		t.Error("all three block scope slices should have length 1 after enter")
	}

	// Simulate OpJSBlockScopeExit.
	vm.jsBlockScopes = vm.jsBlockScopes[:len(vm.jsBlockScopes)-1]
	vm.jsBlockScopeConst = vm.jsBlockScopeConst[:len(vm.jsBlockScopeConst)-1]
	vm.jsBlockScopeTDZ = vm.jsBlockScopeTDZ[:len(vm.jsBlockScopeTDZ)-1]
	vm.jsBlockScopeDepth--

	if len(vm.jsBlockScopes) != 0 || len(vm.jsBlockScopeConst) != 0 || len(vm.jsBlockScopeTDZ) != 0 {
		t.Error("all three block scope slices should be empty after exit")
	}
}

// ─────────────────────────────────────────────────────────────────────────────
// eval / arguments binding restrictions in strict mode
// ─────────────────────────────────────────────────────────────────────────────

// TestJScriptStrictMode_EvalAssignmentError verifies that assigning to "eval"
// in strict mode raises a compile-time SyntaxError.
func TestJScriptStrictMode_EvalAssignmentError(t *testing.T) {
	err := compileJScript2(jscriptSrc(`"use strict"; eval = 5;`))
	if err == nil {
		t.Error("expected compile error for 'eval = 5' in strict mode, got none")
	}
}

// TestJScriptStrictMode_ArgumentsAssignmentError verifies that assigning to
// "arguments" in strict mode raises a compile-time SyntaxError.
func TestJScriptStrictMode_ArgumentsAssignmentError(t *testing.T) {
	err := compileJScript2(jscriptSrc(`"use strict"; arguments = [];`))
	if err == nil {
		t.Error("expected compile error for 'arguments = []' in strict mode, got none")
	}
}

// TestJScriptStrictMode_EvalParameterError verifies that using "eval" as a
// function parameter name in strict mode raises a compile-time SyntaxError.
func TestJScriptStrictMode_EvalParameterError(t *testing.T) {
	err := compileJScript2(jscriptSrc(`"use strict"; function f(eval) { return eval; }`))
	if err == nil {
		t.Error("expected compile error for 'function f(eval)' in strict mode, got none")
	}
}

// TestJScriptStrictMode_ArgumentsParameterError verifies that using "arguments"
// as a function parameter name in strict mode raises a compile-time SyntaxError.
func TestJScriptStrictMode_ArgumentsParameterError(t *testing.T) {
	err := compileJScript2(jscriptSrc(`"use strict"; function f(arguments) { return arguments; }`))
	if err == nil {
		t.Error("expected compile error for 'function f(arguments)' in strict mode, got none")
	}
}

// TestJScriptNonStrict_EvalAssignmentAllowed verifies that assigning to "eval"
// in non-strict mode compiles without error (ES5 legacy behavior).
func TestJScriptNonStrict_EvalAssignmentAllowed(t *testing.T) {
	err := compileJScript2(jscriptSrc(`var eval = 5;`))
	if err != nil {
		t.Errorf("did not expect compile error for 'var eval' in non-strict mode, got: %v", err)
	}
}

// ─────────────────────────────────────────────────────────────────────────────
// delete identifier restriction in strict mode
// ─────────────────────────────────────────────────────────────────────────────

// TestJScriptStrictMode_DeleteIdentifierError verifies that deleting a variable
// identifier in strict mode raises a compile-time SyntaxError.
func TestJScriptStrictMode_DeleteIdentifierError(t *testing.T) {
	err := compileJScript2(jscriptSrc(`"use strict"; var x = 1; delete x;`))
	if err == nil {
		t.Error("expected compile error for 'delete x' in strict mode, got none")
	}
}

// TestJScriptNonStrict_DeleteIdentifierAllowed verifies that deleting a variable
// identifier in non-strict mode compiles without error.
func TestJScriptNonStrict_DeleteIdentifierAllowed(t *testing.T) {
	err := compileJScript2(jscriptSrc(`var x = 1; delete x;`))
	if err != nil {
		t.Errorf("did not expect compile error for 'delete x' in non-strict mode, got: %v", err)
	}
}

// TestJScriptNonStrict_DeletePropertyAlwaysAllowed verifies that deleting an
// object property in strict mode compiles without error.
func TestJScriptNonStrict_DeletePropertyAlwaysAllowed(t *testing.T) {
	err := compileJScript2(jscriptSrc(`"use strict"; var obj = {a: 1}; delete obj.a;`))
	if err != nil {
		t.Errorf("did not expect compile error for 'delete obj.a' in strict mode, got: %v", err)
	}
}

// ─────────────────────────────────────────────────────────────────────────────
// duplicate parameter detection in strict mode
// ─────────────────────────────────────────────────────────────────────────────

// TestJScriptStrictMode_DuplicateParameterError verifies that duplicate
// parameter names in a strict-mode function raise a compile-time SyntaxError.
func TestJScriptStrictMode_DuplicateParameterError(t *testing.T) {
	err := compileJScript2(jscriptSrc(`"use strict"; function f(x, x) { return x; }`))
	if err == nil {
		t.Error("expected compile error for duplicate parameter 'x' in strict mode, got none")
	}
}

// TestJScriptNonStrict_DuplicateParameterAllowed verifies that duplicate
// parameter names in non-strict mode are allowed (ES5 legacy behavior).
func TestJScriptNonStrict_DuplicateParameterAllowed(t *testing.T) {
	err := compileJScript2(jscriptSrc(`function f(x, x) { return x; }`))
	if err != nil {
		t.Errorf("did not expect compile error for duplicate param in non-strict mode, got: %v", err)
	}
}

// ─────────────────────────────────────────────────────────────────────────────
// let block scoping
// ─────────────────────────────────────────────────────────────────────────────

// TestJScriptLetBlockScope verifies that let scopes correctly to a block and
// the value is accessible inside that block.
func TestJScriptLetBlockScope(t *testing.T) {
	src := jscriptSrc(`
		var result = "";
		{ let x = "block"; result = x; }
		Response.Write(result);
	`)
	out, err := runJScript2(t, src)
	if err != nil {
		t.Fatalf("unexpected runtime error: %v", err)
	}
	if out != "block" {
		t.Errorf("expected 'block', got: %q", out)
	}
}

// TestJScriptLetShadowing verifies that let inside a block correctly shadows
// an outer variable with the same name without affecting the outer binding.
func TestJScriptLetShadowing(t *testing.T) {
	src := jscriptSrc(`
		var x = "outer";
		{ let x = "inner"; Response.Write(x + ","); }
		Response.Write(x);
	`)
	out, err := runJScript2(t, src)
	if err != nil {
		t.Fatalf("unexpected runtime error: %v", err)
	}
	if out != "inner,outer" {
		t.Errorf("expected 'inner,outer', got: %q", out)
	}
}

// TestJScriptLetMutability verifies that let-declared variables can be
// reassigned within their block.
func TestJScriptLetMutability(t *testing.T) {
	src := jscriptSrc(`
		{ let x = 1; x = 2; Response.Write(x); }
	`)
	out, err := runJScript2(t, src)
	if err != nil {
		t.Fatalf("unexpected runtime error: %v", err)
	}
	if out != "2" {
		t.Errorf("expected '2', got: %q", out)
	}
}

// ─────────────────────────────────────────────────────────────────────────────
// const block scoping and immutability
// ─────────────────────────────────────────────────────────────────────────────

// TestJScriptConstAccessInsideBlock verifies that a const variable is readable
// within the block where it is declared.
func TestJScriptConstAccessInsideBlock(t *testing.T) {
	src := jscriptSrc(`
		var result = "";
		{ const x = "inner"; result = x; }
		Response.Write(result);
	`)
	out, err := runJScript2(t, src)
	if err != nil {
		t.Fatalf("unexpected runtime error: %v", err)
	}
	if out != "inner" {
		t.Errorf("expected 'inner', got: %q", out)
	}
}

// TestJScriptConstMultipleBlocks verifies that const variables in different
// blocks with the same name are fully independent.
func TestJScriptConstMultipleBlocks(t *testing.T) {
	src := jscriptSrc(`
		var result = "";
		{ const x = "first"; result = result + x; }
		{ const x = "second"; result = result + x; }
		Response.Write(result);
	`)
	out, err := runJScript2(t, src)
	if err != nil {
		t.Fatalf("unexpected runtime error: %v", err)
	}
	if out != "firstsecond" {
		t.Errorf("expected 'firstsecond', got: %q", out)
	}
}

// TestJScriptConstImmutability verifies that reassigning a const variable
// raises a TypeError at runtime, which can be caught by a try/catch.
func TestJScriptConstImmutability(t *testing.T) {
	src := jscriptSrc(`
		var caught = false;
		try {
			const x = 5;
			x = 10;
		} catch(e) {
			caught = true;
		}
		Response.Write(caught ? "TypeError caught" : "no error");
	`)
	out, err := runJScript2(t, src)
	if err != nil {
		t.Logf("runtime error (const TypeError may have bubbled up): %v", err)
	}
	if !strings.Contains(out, "TypeError caught") {
		t.Errorf("expected 'TypeError caught' for const reassignment, got: %q", out)
	}
}

// ─────────────────────────────────────────────────────────────────────────────
// for-loop with let: basic iteration
// ─────────────────────────────────────────────────────────────────────────────

// TestJScriptForLetBasicSum verifies that a for-loop with a let initializer
// correctly accumulates values.
func TestJScriptForLetBasicSum(t *testing.T) {
	src := jscriptSrc(`
		var sum = 0;
		for (let i = 0; i < 5; i++) { sum = sum + i; }
		Response.Write(sum);
	`)
	out, err := runJScript2(t, src)
	if err != nil {
		t.Fatalf("unexpected runtime error: %v", err)
	}
	if out != "10" {
		t.Errorf("expected '10', got: %q", out)
	}
}

// TestJScriptForLetBreak verifies that break inside a for-let loop stops
// iteration at the correct point.
func TestJScriptForLetBreak(t *testing.T) {
	src := jscriptSrc(`
		var count = 0;
		for (let i = 0; i < 10; i++) {
			if (i === 3) break;
			count++;
		}
		Response.Write(count);
	`)
	out, err := runJScript2(t, src)
	if err != nil {
		t.Fatalf("unexpected runtime error: %v", err)
	}
	if out != "3" {
		t.Errorf("expected '3', got: %q", out)
	}
}

// TestJScriptForLetContinue verifies that continue inside a for-let loop
// skips the correct iterations.
func TestJScriptForLetContinue(t *testing.T) {
	src := jscriptSrc(`
		var sum = 0;
		for (let i = 0; i < 5; i++) {
			if (i === 2) continue;
			sum = sum + i;
		}
		Response.Write(sum);
	`)
	out, err := runJScript2(t, src)
	if err != nil {
		t.Fatalf("unexpected runtime error: %v", err)
	}
	// 0 + 1 + (skip 2) + 3 + 4 = 8
	if out != "8" {
		t.Errorf("expected '8', got: %q", out)
	}
}

// ─────────────────────────────────────────────────────────────────────────────
// for-loop with let: per-iteration binding for closures
// ─────────────────────────────────────────────────────────────────────────────

// TestJScriptForLetPerIterationBinding verifies that for-let creates a fresh
// binding per iteration so that closures capture distinct values.
func TestJScriptForLetPerIterationBinding(t *testing.T) {
	src := jscriptSrc(`
		var funcs = [];
		for (let i = 0; i < 3; i++) {
			funcs.push(function() { return i; });
		}
		Response.Write(funcs[0]() + "," + funcs[1]() + "," + funcs[2]());
	`)
	out, err := runJScript2(t, src)
	if err != nil {
		t.Fatalf("unexpected runtime error: %v", err)
	}
	if out != "0,1,2" {
		t.Errorf("expected '0,1,2' for per-iteration let binding, got: %q", out)
	}
}

// TestJScriptForVarClosureCaptureFinalValue verifies that a var-declared loop
// variable causes all closures to capture the final shared value (ES5 legacy
// regression guard).
func TestJScriptForVarClosureCaptureFinalValue(t *testing.T) {
	src := jscriptSrc(`
		var funcs = [];
		for (var i = 0; i < 3; i++) {
			funcs.push(function() { return i; });
		}
		Response.Write(funcs[0]() + "," + funcs[1]() + "," + funcs[2]());
	`)
	out, err := runJScript2(t, src)
	if err != nil {
		t.Fatalf("unexpected runtime error: %v", err)
	}
	// All closures share the same var, which ends up at 3 after the loop.
	if out != "3,3,3" {
		t.Errorf("expected '3,3,3' for var closure capture, got: %q", out)
	}
}

// ─────────────────────────────────────────────────────────────────────────────
// for-in with let
// ─────────────────────────────────────────────────────────────────────────────

// TestJScriptForInLet verifies that for-in with a let declaration correctly
// iterates over an object's enumerable keys.
func TestJScriptForInLet(t *testing.T) {
	src := jscriptSrc(`
		var obj = {a: 1, b: 2};
		var keys = [];
		for (let k in obj) { keys.push(k); }
		keys.sort();
		Response.Write(keys.join(","));
	`)
	out, err := runJScript2(t, src)
	if err != nil {
		t.Fatalf("unexpected runtime error: %v", err)
	}
	if out != "a,b" {
		t.Errorf("expected 'a,b', got: %q", out)
	}
}

// ─────────────────────────────────────────────────────────────────────────────
// TypeError for non-writable properties in strict mode
// ─────────────────────────────────────────────────────────────────────────────

// TestJScriptStrictMode_ReadOnlyPropertyTypeError verifies that assigning to a
// non-writable property in strict mode raises a catchable TypeError.
func TestJScriptStrictMode_ReadOnlyPropertyTypeError(t *testing.T) {
	src := jscriptSrc(`
		"use strict";
		var obj = {};
		Object.defineProperty(obj, "x", {value: 42, writable: false, configurable: false});
		var caught = false;
		try { obj.x = 100; } catch(e) { caught = true; }
		Response.Write(caught ? "TypeError caught" : "no error");
	`)
	out, err := runJScript2(t, src)
	if err != nil {
		t.Logf("runtime error (TypeError may have bubbled up): %v", err)
	}
	if !strings.Contains(out, "TypeError caught") {
		t.Errorf("expected 'TypeError caught' for non-writable in strict mode, got: %q", out)
	}
}

// TestJScriptNonStrict_ReadOnlyPropertySilent verifies that assigning to a
// non-writable property in non-strict mode silently fails (value unchanged).
func TestJScriptNonStrict_ReadOnlyPropertySilent(t *testing.T) {
	src := jscriptSrc(`
		var obj = {};
		Object.defineProperty(obj, "x", {value: 42, writable: false, configurable: false});
		obj.x = 100;
		Response.Write(obj.x);
	`)
	out, err := runJScript2(t, src)
	if err != nil {
		t.Fatalf("unexpected runtime error: %v", err)
	}
	if out != "42" {
		t.Errorf("expected '42' (write silently ignored), got: %q", out)
	}
}

// ─────────────────────────────────────────────────────────────────────────────
// Compiler correctness — structure tests
// ─────────────────────────────────────────────────────────────────────────────

// TestJScriptLexicalDeclarationCompilable verifies that the compiler handles
// a valid let/const block without panicking.
func TestJScriptLexicalDeclarationCompilable(t *testing.T) {
	err := compileJScript2(jscriptSrc(`{ let a = 1; const b = 2; a = 3; Response.Write(a); }`))
	if err != nil {
		t.Errorf("did not expect compile error for valid let/const block, got: %v", err)
	}
}

// TestJScriptStrictMode_LexicalRestrictedName verifies that using "eval" or
// "arguments" as a let/const identifier in strict mode is a SyntaxError.
func TestJScriptStrictMode_LexicalRestrictedName(t *testing.T) {
	err := compileJScript2(jscriptSrc(`"use strict"; { let eval = 5; }`))
	if err == nil {
		t.Error("expected compile error for 'let eval' in strict mode, got none")
	}
}
