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
	ClassDef  *ClassDef
	Variables map[string]interface{}
	Context   *ExecutionContext
}

// NewClassInstance creates a new instance and initializes variables
func NewClassInstance(def *ClassDef, ctx *ExecutionContext) (*ClassInstance, error) {
	inst := &ClassInstance{
		ClassDef:  def,
		Variables: make(map[string]interface{}),
		Context:   ctx,
	}

	// Initialize variables
	for name := range def.Variables {
		inst.Variables[strings.ToLower(name)] = nil
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
	if nameLower == "" && ci.ClassDef != nil && ci.ClassDef.DefaultMethod != "" {
		nameLower = ci.ClassDef.DefaultMethod
	} else if nameLower == "" && ci.ClassDef != nil {
		// Fallback: single-member classes should still allow default-style invocation
		if len(ci.ClassDef.Functions) == 1 {
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

	// 1. Check Public Methods (Sub)
	if sub, ok := ci.ClassDef.Methods[nameLower]; ok {
		return ci.executeMethod(sub, args)
	}

	// 2. Check Public Functions
	if fn, ok := ci.ClassDef.Functions[nameLower]; ok {
		return ci.executeMethod(fn, args)
	}

	// 3. Check Public Property Get (with arguments)
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

	// 1. Check Variables
	if _, ok := ci.ClassDef.Variables[nameLower]; ok {
		val := ci.Variables[nameLower]
		return val, true, nil
	}

	// 2. Check Property Get
	if props, ok := ci.ClassDef.Properties[nameLower]; ok {
		for _, p := range props {
			if p.Type == PropGet {
				val, err := ci.executeMethod(p.Node, []interface{}{})
				return val, true, err
			}
		}
	}

	// 3. Check Functions (0-args)
	if fn, ok := ci.ClassDef.Functions[nameLower]; ok {
		if len(fn.Parameters) == 0 {
			val, err := ci.executeMethod(fn, []interface{}{})
			return val, true, err
		}
	}

	if sub, ok := ci.ClassDef.Methods[nameLower]; ok {
		if len(sub.Parameters) == 0 {
			val, err := ci.executeMethod(sub, []interface{}{})
			return val, true, err
		}
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

	// 6. BIND ARGUMENTS
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

	for _, stmt := range body {
		if err := v.VisitStatement(stmt); err != nil {
			if pe, ok := err.(*ProcedureExitError); ok {
				if expectedExitKind == "" || pe.Kind == expectedExitKind {
					break
				}
				return nil, nil
			}
			return nil, err
		}
	}

	// 9. RETRIEVE RETURN VALUE
	if funcName != "" {
		val, _ := ci.Context.GetVariable(funcName)
		return val, nil
	}

	return nil, nil
}
