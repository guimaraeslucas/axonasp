package server

import (
	"fmt"
	"strings"

	"github.com/guimaraeslucas/vbscript-go/ast"
)

type Visibility int

const (
	VisPublic Visibility = iota
	VisPrivate
)

type ClassMemberVar struct {
	Name       string
	Visibility Visibility
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
	Name          string
	Variables     map[string]ClassMemberVar
	Methods       map[string]*ast.SubDeclaration      // Public Subs
	Functions     map[string]*ast.FunctionDeclaration // Public Functions
	Properties    map[string][]PropertyDef            // Properties can be overloaded by type (Get/Let/Set)
	PrivateMethods map[string]ast.Node                // Private Subs/Functions
}

// ClassInstance represents an instance of a Class
type ClassInstance struct {
	ClassDef  *ClassDef
	Variables map[string]interface{}
	Context   *ExecutionContext
}

// NewClassInstance creates a new instance and initializes variables
func NewClassInstance(def *ClassDef, ctx *ExecutionContext) *ClassInstance {
	inst := &ClassInstance{
		ClassDef:  def,
		Variables: make(map[string]interface{}),
		Context:   ctx,
	}

	// Initialize variables
	for name := range def.Variables {
		inst.Variables[strings.ToLower(name)] = nil // Default to Empty
	}

	// Initialize Class_Initialize if present
	if sub, ok := def.PrivateMethods["class_initialize"]; ok {
		inst.executeMethod(sub, []interface{}{})
	} else if sub, ok := def.Methods["class_initialize"]; ok {
		inst.executeMethod(sub, []interface{}{})
	}

	return inst
}

// GetProperty implements property access
func (ci *ClassInstance) GetProperty(name string) interface{} {
	nameLower := strings.ToLower(name)

	// 1. Try Public Variables
	if vDef, ok := ci.ClassDef.Variables[nameLower]; ok {
		if vDef.Visibility == VisPublic {
			return ci.Variables[nameLower]
		}
	}

	// 2. Try Property Get
	if props, ok := ci.ClassDef.Properties[nameLower]; ok {
		for _, p := range props {
			if p.Type == PropGet && p.Visibility == VisPublic {
				// Execute Property Get
				return ci.executeMethod(p.Node, []interface{}{})
			}
		}
	}

	return nil
}

// SetProperty implements property assignment
func (ci *ClassInstance) SetProperty(name string, value interface{}) error {
	nameLower := strings.ToLower(name)

	// 1. Try Public Variables
	if vDef, ok := ci.ClassDef.Variables[nameLower]; ok {
		if vDef.Visibility == VisPublic {
			ci.Variables[nameLower] = value
			return nil
		}
	}

	// 2. Try Property Let/Set
	// Check if value is an object for PropSet? In VBScript it depends on "Set" keyword in assignment.
	// But here we might rely on the called method type.
	// For now, check both or prioritize based on usage.
	// Usually parser determines if it is a "Set" assignment.
	
	if props, ok := ci.ClassDef.Properties[nameLower]; ok {
		for _, p := range props {
			if (p.Type == PropLet || p.Type == PropSet) && p.Visibility == VisPublic {
				// Execute Property Let/Set
				// We pass the value as the last argument
				_ = ci.executeMethod(p.Node, []interface{}{value})
				return nil
			}
		}
	}

	return fmt.Errorf("property '%s' not found or not writable", name)
}

// CallMethod calls a method on the instance
func (ci *ClassInstance) CallMethod(name string, args ...interface{}) (interface{}, error) {
	nameLower := strings.ToLower(name)

	// 1. Check Public Methods (Sub)
	if sub, ok := ci.ClassDef.Methods[nameLower]; ok {
		return ci.executeMethod(sub, args), nil
	}

	// 2. Check Public Functions
	if fn, ok := ci.ClassDef.Functions[nameLower]; ok {
		return ci.executeMethod(fn, args), nil
	}
	
	// 3. Check Property Get (with arguments)
	if props, ok := ci.ClassDef.Properties[nameLower]; ok {
		for _, p := range props {
			if p.Type == PropGet && p.Visibility == VisPublic {
				return ci.executeMethod(p.Node, args), nil
			}
		}
	}

	return nil, fmt.Errorf("method '%s' not found", name)
}

// executeMethod executes a method or function node
func (ci *ClassInstance) executeMethod(node ast.Node, args []interface{}) interface{} {
	if ci.Context == nil {
		return nil
	}

	// Get executor from Context.Server
	serverObj := ci.Context.Server
	executorInt := serverObj.GetProperty("_executor")
	if executorInt == nil {
		return nil
	}

	executor, ok := executorInt.(*ASPExecutor)
	if !ok {
		return nil
	}

	// Save old context object to restore later (recursion support)
	oldContextObj := ci.Context.GetContextObject()

	// Prepare for execution
	ci.Context.PushScope()
	ci.Context.SetContextObject(ci) // Set 'Me'

	// Defer cleanup
	defer func() {
		ci.Context.SetContextObject(oldContextObj)
		ci.Context.PopScope()
	}()

	// Bind Arguments
	var params []*ast.Parameter
	var body []ast.Statement
	var funcName string

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
	case *ast.PropertyGetDeclaration:
		funcName = n.Identifier.Name
		params = n.Parameters
		body = n.Body
	case *ast.PropertyLetDeclaration:
		params = n.Parameters
		body = n.Body
	case *ast.PropertySetDeclaration:
		params = n.Parameters
		body = n.Body
	default:
		return nil
	}

	// Map args to params
	for i, param := range params {
		if i < len(args) {
			_ = ci.Context.DefineVariable(param.Identifier.Name, args[i])
		} else {
			// Missing argument
			_ = ci.Context.DefineVariable(param.Identifier.Name, nil)
		}
	}

	// Create Visitor
	v := NewASPVisitor(ci.Context, executor)

	// Execute Body
	for _, stmt := range body {
		if err := v.VisitStatement(stmt); err != nil {
			// Check for Exit statements (simplified)
			// In a real implementation we would check the type of error/exit
			// For now, assume any error/return stops execution
			break
		}
	}

	// Return value for Function/Property Get
	if funcName != "" {
		return ci.getReturnValue(funcName)
	}

	return nil
}

func (ci *ClassInstance) getReturnValue(name string) interface{} {
	// Function return value is stored in a variable with the same name
	// inside the local scope (which is currently the top scope)
	val, _ := ci.Context.GetVariable(name)
	return val
}

// GetMember returns a member variable or property (internal access, ignores visibility)
func (ci *ClassInstance) GetMember(name string) (interface{}, bool) {
	nameLower := strings.ToLower(name)
	
	// 1. Check Variables (Private or Public)
	if _, ok := ci.ClassDef.Variables[nameLower]; ok {
		val := ci.Variables[nameLower]
		return val, true
	}
	
	// 2. Check Property Get
	if props, ok := ci.ClassDef.Properties[nameLower]; ok {
		for _, p := range props {
			if p.Type == PropGet {
				// Execute Property Get
				return ci.executeMethod(p.Node, []interface{}{}), true
			}
		}
	}
	
	return nil, false
}

// SetMember sets a member variable or property (internal access)
func (ci *ClassInstance) SetMember(name string, value interface{}) bool {
	nameLower := strings.ToLower(name)
	
	// 1. Check Variables
	if _, ok := ci.ClassDef.Variables[nameLower]; ok {
		ci.Variables[nameLower] = value
		return true
	}
	
	// 2. Try Property Let/Set
	if props, ok := ci.ClassDef.Properties[nameLower]; ok {
		for _, p := range props {
			if p.Type == PropLet || p.Type == PropSet {
				// Execute Property Let/Set
				_ = ci.executeMethod(p.Node, []interface{}{value})
				return true
			}
		}
	}
	
	return false
}
