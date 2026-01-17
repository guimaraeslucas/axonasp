package asp

import (
	"strings"
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
	Params     []string
	LineNum    int
	Visibility Visibility
}

type ClassDef struct {
	Name          string
	Variables     map[string]ClassMemberVar
	Methods       map[string]Procedure     // Subs and Functions
	Properties    map[string][]PropertyDef // Properties can be overloaded by type (Get/Let/Set)
	DefaultMethod string                   // Name of the default method/property
}

// ClassInstance represents an instance of a Class
type ClassInstance struct {
	ClassDef  *ClassDef
	Variables map[string]interface{}
	Ctx       *ExecutionContext // Reference to context for execution
}

// NewClassInstance creates a new instance and initializes variables
func NewClassInstance(def *ClassDef, ctx *ExecutionContext) *ClassInstance {
	inst := &ClassInstance{
		ClassDef:  def,
		Variables: make(map[string]interface{}),
		Ctx:       ctx,
	}

	// Initialize variables (Private and Public)
	for name := range def.Variables {
		inst.Variables[strings.ToLower(name)] = nil // Default to Empty/Conversational nil
	}

	return inst
}

// GetProperty implements Component interface
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
				// We need to invoke the engine to run the property code.
				// Since we are in GetProperty (called by parser/evaluator), we rely on Ctx.Engine
				if ci.Ctx.Engine != nil {
					return ci.Ctx.Engine.ExecuteClassMethod(ci, p.Name, PropGet, []interface{}{}, nil)
				}
			}
		}
	}

	return nil
}

// SetProperty implements Component interface (for simple assignment)
func (ci *ClassInstance) SetProperty(name string, value interface{}) {
	nameLower := strings.ToLower(name)

	// 1. Try Public Variables
	if vDef, ok := ci.ClassDef.Variables[nameLower]; ok {
		if vDef.Visibility == VisPublic {
			ci.Variables[nameLower] = value
			return
		}
	}

	// 2. Try Property Let (or Set if value is object)
	// Distinguish Let/Set based on value type?
	// VBScript uses "Set x.Prop = obj" vs "x.Prop = val".
	// The interpreter parser calls "SetProperty" for both?
	// Actually parser handles "Set" keyword.
	// We need to know if it's Let or Set.
	// For now, try Let.
	propType := PropLet
	if _, ok := value.(Component); ok {
		propType = PropSet
	}

	if props, ok := ci.ClassDef.Properties[nameLower]; ok {
		for _, p := range props {
			if p.Type == propType && p.Visibility == VisPublic {
				if ci.Ctx.Engine != nil {
					ci.Ctx.Engine.ExecuteClassMethod(ci, p.Name, propType, []interface{}{value}, nil)
					return
				}
			}
		}
	}
}

// CallMethod implements Component interface
func (ci *ClassInstance) CallMethod(name string, args []interface{}) interface{} {
	nameLower := strings.ToLower(name)

	// 1. Check Public Methods (Sub/Function)
	if _, ok := ci.ClassDef.Methods[nameLower]; ok {
		// Execute Method
		if ci.Ctx.Engine != nil {
			return ci.Ctx.Engine.ExecuteClassMethod(ci, name, PropGet, args, nil) // PropGet acts as placeholder for "Method"
		}
	}

	// 2. Check Property Get (with arguments)
	if props, ok := ci.ClassDef.Properties[nameLower]; ok {
		for _, p := range props {
			if p.Type == PropGet {
				if ci.Ctx.Engine != nil {
					return ci.Ctx.Engine.ExecuteClassMethod(ci, p.Name, PropGet, args, nil)
				}
			}
		}
	}

	return nil
}
