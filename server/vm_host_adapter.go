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
	"strings"

	"g3pix.com.br/axonasp/experimental"
	"g3pix.com.br/axonasp/vbscript/ast"
)

// VMHostAdapter bridges the experimental VM to the ASP execution context.
type VMHostAdapter struct {
	ctx      *ExecutionContext
	executor *ASPExecutor
}

func NewVMHostAdapter(ctx *ExecutionContext, executor *ASPExecutor) *VMHostAdapter {
	return &VMHostAdapter{
		ctx:      ctx,
		executor: executor,
	}
}

func (h *VMHostAdapter) GetVariable(name string) (interface{}, bool) {
	if h.ctx == nil {
		return nil, false
	}

	nameLower := strings.ToLower(name)
	switch nameLower {
	case "response":
		return h.ctx.Response, true
	case "request":
		return h.ctx.Request, true
	case "session":
		return h.ctx.Session, true
	case "application":
		return h.ctx.Application, true
	case "server":
		return h.ctx.Server, true
	case "err":
		return h.ctx.Err, true
	}

	val, ok := h.ctx.GetVariable(name)
	return val, ok
}

func (h *VMHostAdapter) SetVariable(name string, value interface{}) error {
	if h.ctx == nil {
		return fmt.Errorf("VM host context is nil")
	}
	targetName := name
	if targetName == "" {
		targetName = name
	}
	if decl, ok := value.(*ast.ClassDeclaration); ok {
		visitor := NewASPVisitor(h.ctx, h.executor)
		classDef := visitor.NewClassDefFromDecl(decl)
		if decl != nil && decl.Identifier != nil && decl.Identifier.Name != "" {
			targetName = decl.Identifier.Name
		}
		return h.ctx.SetVariable(targetName, classDef)
	}
	if compiled, ok := value.(*experimental.CompiledClass); ok {
		classDef := &ClassDef{
			Name:            compiled.Name,
			Variables:       make(map[string]ClassMemberVar),
			Methods:         make(map[string]*ast.SubDeclaration),
			Functions:       make(map[string]*ast.FunctionDeclaration),
			Properties:      make(map[string][]PropertyDef),
			PrivateMethods:  make(map[string]ast.Node),
			CompiledMethods: compiled.Methods,
		}
		for _, v := range compiled.Variables {
			classDef.Variables[strings.ToLower(v)] = ClassMemberVar{
				Name:       v,
				Visibility: VisPublic,
			}
		}
		if compiled != nil && compiled.Name != "" {
			targetName = compiled.Name
		}
		return h.ctx.SetVariable(targetName, classDef)
	}
	return h.ctx.SetVariable(targetName, value)
}

func (h *VMHostAdapter) CallFunction(name string, args []interface{}) (interface{}, error) {
	if h.ctx == nil {
		return nil, fmt.Errorf("VM host context is nil")
	}

	nameLower := strings.ToLower(name)
	switch nameLower {
	case "response.write", "responsewrite":
		if len(args) > 0 {
			_ = h.ctx.Response.Write(toString(args[0]))
		}
		return nil, nil
	case "server.createobject", "servercreateobject":
		if len(args) == 0 {
			return nil, fmt.Errorf("CreateObject requires a progID")
		}
		progID := toString(args[0])
		return h.CreateObject(progID)
	}

	if result, handled := evalCustomFunction(nameLower, args, h.ctx); handled {
		return result, nil
	}
	if result, handled := EvalBuiltInFunction(nameLower, args, h.ctx); handled {
		return result, nil
	}

	return nil, fmt.Errorf("function not implemented in VM host: %s", name)
}

func (h *VMHostAdapter) CreateObject(progID string) (interface{}, error) {
	if h.ctx != nil {
		if classVal, ok := h.ctx.GetVariable(progID); ok {
			if classDef, ok := classVal.(*ClassDef); ok {
				return NewClassInstance(classDef, h.ctx)
			}
		}
	}
	if h.executor == nil {
		return nil, fmt.Errorf("executor not available for CreateObject")
	}
	return h.executor.CreateObject(progID)
}

func (h *VMHostAdapter) ExecuteAST(node interface{}) (interface{}, error) {
	if h.ctx == nil || h.executor == nil {
		return nil, fmt.Errorf("VM host context or executor is nil")
	}

	visitor := NewASPVisitor(h.ctx, h.executor)
	execStatement := func(stmt ast.Statement) error {
		if stmt == nil {
			return nil
		}
		if h.ctx.ShouldStop() {
			return fmt.Errorf("RESPONSE_END")
		}
		return visitor.VisitStatement(stmt)
	}

	if program, ok := node.(*ast.Program); ok {
		h.executor.hoistDeclarations(visitor, program)
		for _, stmt := range program.Body {
			if err := execStatement(stmt); err != nil {
				return nil, err
			}
		}
		return nil, nil
	}

	if list, ok := node.(*ast.StatementList); ok {
		for _, stmt := range list.Statements {
			if err := execStatement(stmt); err != nil {
				return nil, err
			}
		}
		return nil, nil
	}

	if stmt, ok := node.(ast.Statement); ok {
		if err := execStatement(stmt); err != nil {
			return nil, err
		}
		return nil, nil
	}

	if expr, ok := node.(ast.Expression); ok {
		return visitor.visitExpression(expr)
	}

	return nil, fmt.Errorf("invalid node type for ExecuteAST: %T", node)
}

func (h *VMHostAdapter) SetIndexed(obj interface{}, indexes []interface{}, value interface{}) error {
	if h.ctx == nil || h.executor == nil {
		return fmt.Errorf("VM host context or executor is nil")
	}

	// Handle native VBArray
	if arrObj, ok := toVBArray(obj); ok {
		if len(indexes) == 1 {
			index := toInt(indexes[0])
			if arrObj.Set(index, value) {
				return nil
			}
			return fmt.Errorf("subscript out of range")
		} else if len(indexes) > 1 {
			current := arrObj
			for i := 0; i < len(indexes)-1; i++ {
				index := toInt(indexes[i])
				inner, exists := current.Get(index)
				if !exists {
					return fmt.Errorf("subscript out of range")
				}
				innerArr, ok := toVBArray(inner)
				if !ok {
					return fmt.Errorf("subscript out of range")
				}
				current = innerArr
			}
			lastIndex := toInt(indexes[len(indexes)-1])
			if current.Set(lastIndex, value) {
				return nil
			}
			return fmt.Errorf("subscript out of range")
		}
	}

	// Handle Map
	if mapObj, ok := obj.(map[string]interface{}); ok {
		if len(indexes) > 0 {
			key := fmt.Sprintf("%v", indexes[0])
			mapObj[key] = value
			return nil
		}
	}

	// Handle Session
	if sessionObj, ok := obj.(*SessionObject); ok {
		if len(indexes) > 0 {
			return sessionObj.SetIndex(indexes[0], value)
		}
	}

	// Handle Application
	if appObj, ok := obj.(*ApplicationObject); ok {
		if len(indexes) > 0 {
			appObj.Set(fmt.Sprintf("%v", indexes[0]), value)
			return nil
		}
	}

	// Handle generic object with SetProperty (for dictionaries, etc.)
	if lib, ok := obj.(interface {
		SetProperty(string, interface{}) error
	}); ok && len(indexes) > 0 {
		key := fmt.Sprintf("%v", indexes[0])
		return lib.SetProperty(key, value)
	}

	return fmt.Errorf("object does not support indexed assignment: %T", obj)
}

var _ experimental.HostEnvironment = (*VMHostAdapter)(nil)
