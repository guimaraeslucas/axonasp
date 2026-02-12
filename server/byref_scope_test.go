package server

import (
	"testing"
	"time"

	"g3pix.com.br/axonasp/vbscript/ast"
)

// TestSetVariableInParentScope_DoesNotCreateNewEntries verifies that
// SetVariableInParentScope does NOT create new entries in the parent scope
// when the variable doesn't already exist there. Instead, it should fall
// through to the context object (class member) or global variables.
//
// This is the root cause fix for the QuickerSite constants caching bug
// where ByRef writeback from isNumeriek(iId) created a spurious "iid"
// entry in the caller's scope, shadowing the class member variable.
func TestSetVariableInParentScope_DoesNotCreateNewEntries(t *testing.T) {
	ctx := NewExecutionContext(nil, nil, "", 30*time.Second)

	// Push two scopes: parent (caller) and current (function)
	ctx.PushScope()
	ctx.PushScope()

	// Parent scope (index 0) has a variable "x" but NOT "y"
	ctx.scopeStack[0]["x"] = 10

	// SetVariableInParentScope for existing variable "x" should update it
	err := ctx.SetVariableInParentScope("x", 20)
	if err != nil {
		t.Fatalf("SetVariableInParentScope(x) failed: %v", err)
	}
	if ctx.scopeStack[0]["x"] != 20 {
		t.Fatalf("expected x=20 in parent scope, got %v", ctx.scopeStack[0]["x"])
	}

	// SetVariableInParentScope for NON-existing variable "y" should NOT create it in parent scope
	err = ctx.SetVariableInParentScope("y", 99)
	if err != nil {
		t.Fatalf("SetVariableInParentScope(y) failed: %v", err)
	}
	if _, exists := ctx.scopeStack[0]["y"]; exists {
		t.Fatalf("expected 'y' NOT to be created in parent scope, but it was! value=%v", ctx.scopeStack[0]["y"])
	}
	// "y" should have gone to global variables instead
	if val, exists := ctx.GetVariable("y"); !exists || val != 99 {
		t.Fatalf("expected y=99 in global variables, got val=%v, exists=%v", val, exists)
	}

	ctx.PopScope()
	ctx.PopScope()
}

// TestSetVariableInParentScope_ClassMemberNotShadowed verifies that
// ByRef writeback for a class member variable does NOT create a local
// copy in the caller's scope, but instead correctly updates the class member.
func TestSetVariableInParentScope_ClassMemberNotShadowed(t *testing.T) {
	ctx := NewExecutionContext(nil, nil, "", 30*time.Second)

	// Create a mock class with a member variable "iid"
	classDef := &ClassDef{
		Name:           "TestClass",
		Variables:      map[string]ClassMemberVar{"iid": {Visibility: VisPublic}},
		Methods:        make(map[string]*ast.SubDeclaration),
		Functions:      make(map[string]*ast.FunctionDeclaration),
		Properties:     make(map[string][]PropertyDef),
		PrivateMethods: make(map[string]ast.Node),
	}

	inst, err := NewClassInstance(classDef, ctx)
	if err != nil {
		t.Fatalf("NewClassInstance failed: %v", err)
	}
	inst.Variables["iid"] = 73

	// Set context object to the class instance (simulating being inside a class method)
	ctx.SetContextObject(inst)

	// Push two scopes: caller scope (e.g., constants()) and function scope (e.g., isNumeriek())
	ctx.PushScope()
	ctx.DefineVariable("sql", "select *") // some local variable in caller scope
	ctx.DefineVariable("rs", nil)         // another local
	ctx.PushScope()                       // function scope
	ctx.DefineVariable("p", 73)           // function parameter

	// Simulate ByRef writeback: SetVariableInParentScope("iid", 73)
	// This should go to the class member, NOT create "iid" in the caller scope
	err = ctx.SetVariableInParentScope("iid", 42)
	if err != nil {
		t.Fatalf("SetVariableInParentScope(iid) failed: %v", err)
	}

	// Verify: "iid" should NOT be in the caller scope (scope[0])
	if _, exists := ctx.scopeStack[0]["iid"]; exists {
		t.Fatalf("REGRESSION: 'iid' was created in caller scope, shadowing class member!")
	}

	// Verify: class member should be updated
	if inst.Variables["iid"] != 42 {
		t.Fatalf("expected class member iid=42, got %v", inst.Variables["iid"])
	}

	ctx.PopScope()
	ctx.PopScope()
	ctx.SetContextObject(nil)
}

// TestSetVariableInParentScope_DeepScopeStack verifies behavior with
// multiple nested scopes — the variable should be found in the correct
// parent scope even with many levels of nesting.
func TestSetVariableInParentScope_DeepScopeStack(t *testing.T) {
	ctx := NewExecutionContext(nil, nil, "", 30*time.Second)

	// Push 4 scopes simulating deep call chain
	ctx.PushScope()
	ctx.scopeStack[0]["globalvar"] = "original"
	ctx.PushScope()
	ctx.PushScope()
	ctx.PushScope() // current function scope

	// SetVariableInParentScope should find "globalvar" in scope[0]
	err := ctx.SetVariableInParentScope("globalvar", "updated")
	if err != nil {
		t.Fatalf("SetVariableInParentScope failed: %v", err)
	}
	if ctx.scopeStack[0]["globalvar"] != "updated" {
		t.Fatalf("expected globalvar='updated' in scope[0], got %v", ctx.scopeStack[0]["globalvar"])
	}

	// Variables not in any scope should go to globals, not pollute parent scopes
	err = ctx.SetVariableInParentScope("newvar", "value")
	if err != nil {
		t.Fatalf("SetVariableInParentScope(newvar) failed: %v", err)
	}
	for i := 0; i < len(ctx.scopeStack); i++ {
		if _, exists := ctx.scopeStack[i]["newvar"]; exists {
			t.Fatalf("'newvar' should not be in scope[%d]", i)
		}
	}
	// Should be in global variables
	ctx.mu.RLock()
	if val, exists := ctx.variables["newvar"]; !exists || val != "value" {
		t.Fatalf("expected newvar='value' in globals, got %v", val)
	}
	ctx.mu.RUnlock()

	ctx.PopScope()
	ctx.PopScope()
	ctx.PopScope()
	ctx.PopScope()
}
