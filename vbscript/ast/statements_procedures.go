/*
 * AxonASP Server - Version 1.0
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
 * ----------------------------------------------------------------------------
 * THIRD PARTY ATTRIBUTION / ORIGINAL SOURCE
 * ----------------------------------------------------------------------------
 * This code was adapted from https://github.com/kmvi/vbscript-parser/
 * ensuring compatibility with VBScript language specifications.
 *
 * Original Copyright (c) [ANO] kmvi (and/or original authors).
 * Licensed under the BSD 3-Clause License.
 *
 * Redistribution and use in source and binary forms, with or without
 * modification, are permitted provided that the following conditions are met:
 *
 * 1. Redistributions of source code must retain the above copyright notice,
 * this list of conditions and the following disclaimer.
 *
 * 2. Redistributions in binary form must reproduce the above copyright notice,
 * this list of conditions and the following disclaimer in the documentation
 * and/or other materials provided with the distribution.
 *
 * 3. Neither the name of the copyright holder nor the names of its
 * contributors may be used to endorse or promote products derived from
 * this software without specific prior written permission.
 *
 * THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS"
 * AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE
 * IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE
 * ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE
 * LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR
 * CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF
 * SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS
 * INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN
 * CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE)
 * ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE
 * POSSIBILITY OF SUCH DAMAGE.
 */
package ast

// ProcedureDeclaration is the base for Sub and Function declarations
type ProcedureDeclaration interface {
	Statement
	isProcedureDeclaration()
}

// BaseProcedureDeclaration provides common functionality
type BaseProcedureDeclaration struct {
	BaseStatement
	AccessModifier MethodAccessModifier
	Identifier     *Identifier
	Parameters     []*Parameter
	Body           Statement
}

func (p *BaseProcedureDeclaration) isProcedureDeclaration() {}

// SubDeclaration represents a Sub declaration
type SubDeclaration struct {
	BaseProcedureDeclaration
}

// NewSubDeclaration creates a new SubDeclaration
func NewSubDeclaration(modifier MethodAccessModifier, id *Identifier, body Statement) *SubDeclaration {
	if id == nil {
		panic("identifier cannot be nil")
	}
	if body == nil {
		panic("body cannot be nil")
	}
	return &SubDeclaration{
		BaseProcedureDeclaration: BaseProcedureDeclaration{
			AccessModifier: modifier,
			Identifier:     id,
			Parameters:     []*Parameter{},
			Body:           body,
		},
	}
}

// InitializeSubDeclaration represents a Class_Initialize sub
type InitializeSubDeclaration struct {
	SubDeclaration
}

const InitializeSubName = "Class_Initialize"

// NewInitializeSubDeclaration creates a new InitializeSubDeclaration
func NewInitializeSubDeclaration(modifier MethodAccessModifier, body Statement) *InitializeSubDeclaration {
	if body == nil {
		panic("body cannot be nil")
	}
	return &InitializeSubDeclaration{
		SubDeclaration: SubDeclaration{
			BaseProcedureDeclaration: BaseProcedureDeclaration{
				AccessModifier: modifier,
				Identifier:     NewIdentifier(InitializeSubName),
				Parameters:     []*Parameter{},
				Body:           body,
			},
		},
	}
}

// TerminateSubDeclaration represents a Class_Terminate sub
type TerminateSubDeclaration struct {
	SubDeclaration
}

const TerminateSubName = "Class_Terminate"

// NewTerminateSubDeclaration creates a new TerminateSubDeclaration
func NewTerminateSubDeclaration(modifier MethodAccessModifier, body Statement) *TerminateSubDeclaration {
	if body == nil {
		panic("body cannot be nil")
	}
	return &TerminateSubDeclaration{
		SubDeclaration: SubDeclaration{
			BaseProcedureDeclaration: BaseProcedureDeclaration{
				AccessModifier: modifier,
				Identifier:     NewIdentifier(TerminateSubName),
				Parameters:     []*Parameter{},
				Body:           body,
			},
		},
	}
}

// FunctionDeclaration represents a Function declaration
type FunctionDeclaration struct {
	BaseProcedureDeclaration
}

// NewFunctionDeclaration creates a new FunctionDeclaration
func NewFunctionDeclaration(modifier MethodAccessModifier, id *Identifier, body Statement) *FunctionDeclaration {
	if id == nil {
		panic("identifier cannot be nil")
	}
	if body == nil {
		panic("body cannot be nil")
	}
	return &FunctionDeclaration{
		BaseProcedureDeclaration: BaseProcedureDeclaration{
			AccessModifier: modifier,
			Identifier:     id,
			Parameters:     []*Parameter{},
			Body:           body,
		},
	}
}

// PropertyDeclaration is the base for property declarations
type PropertyDeclaration interface {
	Statement
	isPropertyDeclaration()
}

// BasePropertyDeclaration provides common functionality for properties
type BasePropertyDeclaration struct {
	BaseStatement
	AccessModifier MethodAccessModifier
	Identifier     *Identifier
	Parameters     []*Parameter
	Body           []Statement
}

func (p *BasePropertyDeclaration) isPropertyDeclaration() {}

// PropertyGetDeclaration represents a Property Get declaration
type PropertyGetDeclaration struct {
	BasePropertyDeclaration
}

// NewPropertyGetDeclaration creates a new PropertyGetDeclaration
func NewPropertyGetDeclaration(modifier MethodAccessModifier, id *Identifier) *PropertyGetDeclaration {
	if id == nil {
		panic("identifier cannot be nil")
	}
	return &PropertyGetDeclaration{
		BasePropertyDeclaration: BasePropertyDeclaration{
			AccessModifier: modifier,
			Identifier:     id,
			Parameters:     []*Parameter{},
			Body:           []Statement{},
		},
	}
}

// PropertySetDeclaration represents a Property Set declaration
type PropertySetDeclaration struct {
	BasePropertyDeclaration
}

// NewPropertySetDeclaration creates a new PropertySetDeclaration
func NewPropertySetDeclaration(modifier MethodAccessModifier, id *Identifier) *PropertySetDeclaration {
	if id == nil {
		panic("identifier cannot be nil")
	}
	return &PropertySetDeclaration{
		BasePropertyDeclaration: BasePropertyDeclaration{
			AccessModifier: modifier,
			Identifier:     id,
			Parameters:     []*Parameter{},
			Body:           []Statement{},
		},
	}
}

// PropertyLetDeclaration represents a Property Let declaration
type PropertyLetDeclaration struct {
	BasePropertyDeclaration
}

// NewPropertyLetDeclaration creates a new PropertyLetDeclaration
func NewPropertyLetDeclaration(modifier MethodAccessModifier, id *Identifier) *PropertyLetDeclaration {
	if id == nil {
		panic("identifier cannot be nil")
	}
	return &PropertyLetDeclaration{
		BasePropertyDeclaration: BasePropertyDeclaration{
			AccessModifier: modifier,
			Identifier:     id,
			Parameters:     []*Parameter{},
			Body:           []Statement{},
		},
	}
}

// ClassDeclaration represents a Class declaration
type ClassDeclaration struct {
	BaseStatement
	Identifier *Identifier
	Members    []Statement
}

// NewClassDeclaration creates a new ClassDeclaration
func NewClassDeclaration(id *Identifier) *ClassDeclaration {
	if id == nil {
		panic("identifier cannot be nil")
	}
	return &ClassDeclaration{
		Identifier: id,
		Members:    []Statement{},
	}
}

// AddMember adds a member to the class
func (c *ClassDeclaration) AddMember(stmt Statement) {
	if stmt != nil {
		c.Members = append(c.Members, stmt)
	}
}
