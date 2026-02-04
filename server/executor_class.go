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
package server

import (
	"fmt"
	//"runtime"
	"strings"

	"g3pix.com.br/axonasp/vbscript/ast"
)

type Visibility int

const (
	VisPublic Visibility = iota
	VisPrivate
)

type ClassMemberVar struct {
	Name       string
	Visibility Visibility
	Dims       []int
	IsDynamic  bool
}

type PropertyType int

const (
	PropGet PropertyType = iota
	PropLet
	PropSet
)

type PropertyDef struct {
	Name       string
	Type       PropertyType
	Node       ast.Node // FunctionDeclaration or SubDeclaration
	Visibility Visibility
}

type ClassDef struct {
	Name           string
	Variables      map[string]ClassMemberVar
	Methods        map[string]*ast.SubDeclaration      // Public Subs
	Functions      map[string]*ast.FunctionDeclaration // Public Functions
	Properties     map[string][]PropertyDef            // Properties can be overloaded by type (Get/Let/Set)
	PrivateMethods map[string]ast.Node                 // Private Subs/Functions
	DefaultMethod  string                              // Lowercase name of the default member
}

// ClassInstance represents an instance of a Class
type ClassInstance struct {
	ClassDef       *ClassDef
	Variables      map[string]interface{}
	Context        *ExecutionContext
	executingProps map[string]bool // Track properties currently being executed to prevent infinite recursion
}

// makeNestedArrayWithBase mirrors ASPVisitor.makeNestedArray but accepts an explicit base.
func makeNestedArrayWithBase(base int, dims []int) *VBArray {
	if len(dims) == 0 {
		return NewVBArray(base, 0)
	}

	lower := base
	upper := dims[0]
	size := upper - lower + 1
	if size < 0 {
		size = 0
	}

	arr := NewVBArray(lower, size)

	if len(dims) > 1 {
		innerDims := dims[1:]
		for i := 0; i < size; i++ {
			arr.Values[i] = makeNestedArrayWithBase(base, innerDims)
		}
	}

	return arr
}

// NewClassInstance creates a new instance and initializes variables
func NewClassInstance(def *ClassDef, ctx *ExecutionContext) (*ClassInstance, error) {
	inst := &ClassInstance{
		ClassDef:       def,
		Variables:      make(map[string]interface{}),
		Context:        ctx,
		executingProps: make(map[string]bool),
	}

	// Initialize variables
	for name, vdef := range def.Variables {
		key := strings.ToLower(name)
		if len(vdef.Dims) > 0 {
			inst.Variables[key] = makeNestedArrayWithBase(ctx.OptionBase(), vdef.Dims)
		} else {
			inst.Variables[key] = nil
		}
	}

	// Initialize Class_Initialize if present
	if sub, ok := def.PrivateMethods["class_initialize"]; ok {
		if _, err := inst.executeMethod(sub, []interface{}{}); err != nil {
			return nil, err
		}
	} else if sub, ok := def.Methods["class_initialize"]; ok {
		if _, err := inst.executeMethod(sub, []interface{}{}); err != nil {
			return nil, err
		}
	}

	return inst, nil
}

// GetName returns the class name
func (ci *ClassInstance) GetName() string {
	return ci.ClassDef.Name
}

// GetProperty implements property access (External Access)
func (ci *ClassInstance) GetProperty(name string) (interface{}, error) {
	nameLower := strings.ToLower(name)

	// 1. Try Public Variables
	if vDef, ok := ci.ClassDef.Variables[nameLower]; ok {
		if vDef.Visibility == VisPublic {
			return ci.Variables[nameLower], nil
		}
	}

	// 2. Try Public Property Get
	if props, ok := ci.ClassDef.Properties[nameLower]; ok {
		for _, p := range props {
			if p.Type == PropGet && p.Visibility == VisPublic {
				return ci.executeMethod(p.Node, []interface{}{})
			}
		}
	}

	return nil, nil // Not found is not strictly an error in GetProperty interface? Or is it?
	// Usually GetProperty returns value. If not found, nil?
	// But if execution failed, we must return error.
	// The interface in executor.go for GetProperty is: GetProperty(string) interface{}
	// It doesn't support error return yet. We will need to update that too if we want full safety.
	// But `GetMember` (Internal) is used for variable lookup.
	// `GetProperty` is used for `obj.Prop`.
	// For now, I will return (interface{}, error) here and update caller to handle it.
}

// SetProperty implements property assignment (External Access)
func (ci *ClassInstance) SetProperty(name string, value interface{}) error {
	nameLower := strings.ToLower(name)

	// 1. Try Public Variables
	if vDef, ok := ci.ClassDef.Variables[nameLower]; ok {
		if vDef.Visibility == VisPublic {
			ci.Variables[nameLower] = value
			return nil
		}
	}

	// 2. Try Public Property Let/Set
	if props, ok := ci.ClassDef.Properties[nameLower]; ok {
		for _, p := range props {
			if (p.Type == PropLet || p.Type == PropSet) && p.Visibility == VisPublic {
				_, err := ci.executeMethod(p.Node, []interface{}{value})
				return err
			}
		}
	}

	return fmt.Errorf("property '%s' not found or not writable", name)
}

// CallMethod calls a method on the instance (External Access)
func (ci *ClassInstance) CallMethod(name string, args ...interface{}) (interface{}, error) {
	nameLower := strings.ToLower(name)
	// Debug logging commented out for performance
	// if strings.ToLower(ci.ClassDef.Name) == "cls_asplite" && nameLower == "json" {
	// 	//fmt.Printf("[DEBUG] cls_asplite.json called from CallMethod (stack trace needed) - this should NOT happen repeatedly\n")
	// 	// Print a mini stack trace
	// 	for i := 0; i < 10; i++ {
	// 		_, file, line, ok := runtime.Caller(i)
	// 		if !ok {
	// 			break
	// 		}
	// 		_ = line
	// 		if strings.Contains(file, "axonasp") {
	// 			//fmt.Printf("[DEBUG]   at %s:%d\n", file, line)
	// 		}
	// 	}
	// }
	if nameLower == "" && ci.ClassDef != nil {
		// Prefer the explicitly declared default member when available
		if ci.ClassDef.DefaultMethod != "" {
			nameLower = ci.ClassDef.DefaultMethod
		} else {
			// Graceful fallback: allow direct invocation to hit a common Exec entrypoint when no default was recorded
			if _, ok := ci.ClassDef.Methods["exec"]; ok {
				nameLower = "exec"
			} else if _, ok := ci.ClassDef.Functions["exec"]; ok {
				nameLower = "exec"
			} else if len(ci.ClassDef.Functions) == 1 {
				// Single-member classes should still allow default-style invocation
				for k := range ci.ClassDef.Functions {
					nameLower = k
				}
			} else if len(ci.ClassDef.Methods) == 1 {
				for k := range ci.ClassDef.Methods {
					nameLower = k
				}
			} else if len(ci.ClassDef.Properties) == 1 {
				for k := range ci.ClassDef.Properties {
					nameLower = k
				}
			}
		}
	}

	// 1. Check Public Methods (Sub)
	if sub, ok := ci.ClassDef.Methods[nameLower]; ok {
		return ci.executeMethod(sub, args)
	}

	// 2. Check Public Functions
	if fn, ok := ci.ClassDef.Functions[nameLower]; ok {
		result, err := ci.executeMethod(fn, args)
		if nameLower == "json" {
			//fmt.Printf("[DEBUG] CallMethod(json): executeMethod returned, result type=%T, err=%v\n", result, err)
		}
		return result, err
	}

	// 3. Check Private Methods (Sub/Function) - can be called from within class context
	if node, ok := ci.ClassDef.PrivateMethods[nameLower]; ok {
		return ci.executeMethod(node, args)
	}

	// 4. Check Public Property Get (with arguments)
	if props, ok := ci.ClassDef.Properties[nameLower]; ok {
		for _, p := range props {
			if p.Type == PropGet && p.Visibility == VisPublic {
				return ci.executeMethod(p.Node, args)
			}
		}
	}

	return nil, fmt.Errorf("method '%s' not found", name)
}

// GetMember returns a member variable or property (Internal Access)
// Returns: value, found, error
func (ci *ClassInstance) GetMember(name string) (interface{}, bool, error) {
	nameLower := strings.ToLower(name)

	if nameLower == "json" || nameLower == "p_json" {
		//fmt.Printf("[DEBUG] GetMember(%s) called on %s\n", name, ci.ClassDef.Name)
	}

	// 1. Check Variables
	if _, ok := ci.ClassDef.Variables[nameLower]; ok {
		val := ci.Variables[nameLower]
		if nameLower == "json" || nameLower == "p_json" {
			//fmt.Printf("[DEBUG] GetMember(%s): found in Variables, val=%T\n", name, val)
		}
		return val, true, nil
	}

	// 2. Check Property Get - with recursion protection
	if props, ok := ci.ClassDef.Properties[nameLower]; ok {
		for _, p := range props {
			if p.Type == PropGet {
				// Prevent infinite recursion by checking if we're already executing this property
				if ci.executingProps == nil {
					ci.executingProps = make(map[string]bool)
				}
				if ci.executingProps[nameLower] {
					// Already executing this property, return not found to avoid recursion
					return nil, false, nil
				}
				ci.executingProps[nameLower] = true
				val, err := ci.executeMethod(p.Node, []interface{}{})
				delete(ci.executingProps, nameLower)
				return val, true, err
			}
		}
	}

	// NOTE: We deliberately do NOT auto-execute Functions or Methods here.
	// VBScript functions should only be called when explicitly invoked with parentheses.
	// This prevents infinite recursion when a function body assigns to the function name:
	//   Function Foo()
	//       Foo = "value"  ' <-- This should NOT trigger Foo() again!
	//   End Function
	// The caller (resolveCall) will handle explicit function invocation.

	// 3. Check if it's a function/method (return false to indicate not a variable)
	// The caller can then decide whether to invoke it
	if _, ok := ci.ClassDef.Functions[nameLower]; ok {
		return nil, false, nil
	}

	if _, ok := ci.ClassDef.Methods[nameLower]; ok {
		return nil, false, nil
	}

	// 4. Check Private Methods (Subs/Functions)
	if node, ok := ci.ClassDef.PrivateMethods[nameLower]; ok {
		// Private methods are stored as generic Nodes (Sub/Function/Property)
		// Return the AST Node itself to let resolveCall handle it
		return node, true, nil
	}

	return nil, false, nil
}

// SetMember sets a member variable or property (Internal Access)
// Returns: handled, error
func (ci *ClassInstance) SetMember(name string, value interface{}) (bool, error) {
	nameLower := strings.ToLower(name)

	// 1. Check Variables
	if _, ok := ci.ClassDef.Variables[nameLower]; ok {
		ci.Variables[nameLower] = value
		return true, nil
	}

	// 2. Try Property Let/Set
	if props, ok := ci.ClassDef.Properties[nameLower]; ok {
		for _, p := range props {
			if p.Type == PropLet || p.Type == PropSet {
				_, err := ci.executeMethod(p.Node, []interface{}{value})
				return true, err
			}
		}
	}

	return false, nil
}

// executeMethod executes a method or function node within the class context
func (ci *ClassInstance) executeMethod(node ast.Node, args []interface{}) (interface{}, error) {
	// Debug: trace all executeMethod calls (commented out for performance)
	// nodeName := "unknown"
	// if fn, ok := node.(*ast.FunctionDeclaration); ok && fn.Identifier != nil {
	// 	nodeName = fn.Identifier.Name
	// } else if sub, ok := node.(*ast.SubDeclaration); ok && sub.Identifier != nil {
	// 	nodeName = sub.Identifier.Name
	// }
	//fmt.Printf("[DEBUG] executeMethod ENTRY: class=%s, method=%s, args=%d\n", ci.ClassDef.Name, nodeName, len(args))

	if ci.Context == nil {
		return nil, fmt.Errorf("no context")
	}

	// Get executor
	serverObj := ci.Context.Server
	if serverObj == nil {
		return nil, fmt.Errorf("no executor")
	}

	executorInt := serverObj.GetExecutor()
	if executorInt == nil {
		return nil, fmt.Errorf("no executor")
	}

	executor, ok := executorInt.(*ASPExecutor)
	if !ok {
		return nil, fmt.Errorf("invalid executor")
	}

	// 1. SAVE PREVIOUS CONTEXT
	oldContextObj := ci.Context.GetContextObject()

	// 2. PUSH NEW SCOPE
	ci.Context.PushScope()

	// 3. SET 'ME' CONTEXT
	ci.Context.SetContextObject(ci)

	// 4. DEFER CLEANUP
	defer func() {
		ci.Context.SetContextObject(oldContextObj)
		ci.Context.PopScope()
	}()

	// 5. IDENTIFY METHOD SIGNATURE
	var params []*ast.Parameter
	var body []ast.Statement
	var funcName string
	var expectedExitKind string

	switch n := node.(type) {
	case *ast.SubDeclaration:
		params = n.Parameters
		if n.Body != nil {
			if list, ok := n.Body.(*ast.StatementList); ok {
				body = list.Statements
			} else {
				body = []ast.Statement{n.Body}
			}
		}
		expectedExitKind = "sub"
	case *ast.FunctionDeclaration:
		funcName = n.Identifier.Name
		params = n.Parameters
		if n.Body != nil {
			if list, ok := n.Body.(*ast.StatementList); ok {
				body = list.Statements
			} else {
				body = []ast.Statement{n.Body}
			}
		}
		expectedExitKind = "function"
	case *ast.PropertyGetDeclaration:
		funcName = n.Identifier.Name
		params = n.Parameters
		body = n.Body
		expectedExitKind = "property"
	case *ast.PropertyLetDeclaration:
		params = n.Parameters
		body = n.Body
		expectedExitKind = "property"
	case *ast.PropertySetDeclaration:
		params = n.Parameters
		body = n.Body
		expectedExitKind = "property"
	default:
		return nil, fmt.Errorf("invalid method node")
	}

	// 6. BIND ARGUMENTS with ByRef support
	// Note: ByRef for class methods is now handled in callClassMethodWithRefs
	for i, param := range params {
		paramName := param.Identifier.Name
		var val interface{}
		if i < len(args) {
			val = args[i]
		}
		_ = ci.Context.DefineVariable(paramName, val)
	}

	// 7. DEFINE RETURN VARIABLE
	if funcName != "" {
		_ = ci.Context.DefineVariable(funcName, nil)
	}

	// 8. EXECUTE BODY
	v := NewASPVisitor(ci.Context, executor)
	// Set currentFunctionName so that return value assignment (e.g., "json = value") works correctly
	// This prevents infinite recursion when the function name is used inside the function body
	v.currentFunctionName = funcName
	if strings.ToLower(funcName) == "json" {
		//fmt.Printf("[DEBUG] executeMethod for 'json' function, currentFunctionName set to: %q, body has %d statements\n", funcName, len(body))
	}

	for _, stmt := range body {
		if strings.ToLower(funcName) == "json" {
			//fmt.Printf("[DEBUG] json function: executing statement %d/%d: %T\n", i+1, len(body), stmt)
		}
		if err := v.VisitStatement(stmt); err != nil {
			if pe, ok := err.(*ProcedureExitError); ok {
				if expectedExitKind == "" || pe.Kind == expectedExitKind {
					break
				}
				return nil, nil
			}
			return nil, err
		}
		if strings.ToLower(funcName) == "json" {
			//fmt.Printf("[DEBUG] json function: statement %d completed\n", i+1)
		}
	}

	if strings.ToLower(funcName) == "json" {
		//fmt.Printf("[DEBUG] json function: all statements executed, retrieving return value\n")
	}

	// 9. RETRIEVE RETURN VALUE
	if funcName != "" {
		val, _ := ci.Context.GetVariable(funcName)
		return val, nil
	}

	return nil, nil
}
