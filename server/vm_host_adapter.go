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
	val, ok := h.ctx.GetVariable(name)
	return val, ok
}

func (h *VMHostAdapter) SetVariable(name string, value interface{}) error {
	if h.ctx == nil {
		return fmt.Errorf("VM host context is nil")
	}
	if decl, ok := value.(*ast.ClassDeclaration); ok {
		visitor := NewASPVisitor(h.ctx, h.executor)
		classDef := visitor.NewClassDefFromDecl(decl)
		return h.ctx.SetVariable(decl.Identifier.Name, classDef)
	}
	return h.ctx.SetVariable(name, value)
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

var _ experimental.HostEnvironment = (*VMHostAdapter)(nil)
