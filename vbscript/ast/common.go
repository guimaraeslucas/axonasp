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

// Program represents the root node of a VBScript AST
type Program struct {
	BaseNode
	OptionExplicit bool
	OptionCompare  OptionCompareMode
	OptionBase     int
	Body           []Statement
	Comments       []*Comment
}

// NewProgram creates a new Program node
func NewProgram(optionExplicit bool, optionCompare OptionCompareMode, optionBase int) *Program {
	return &Program{
		OptionExplicit: optionExplicit,
		OptionCompare:  optionCompare,
		OptionBase:     optionBase,
		Body:           []Statement{},
		Comments:       []*Comment{},
	}
}

// OptionCompareMode defines how string comparisons are evaluated
type OptionCompareMode int

const (
	OptionCompareBinary OptionCompareMode = iota
	OptionCompareText
)

func (p *Program) isStatement() {}

// CommentType represents the type of comment
type CommentType int

const (
	CommentTypeRem CommentType = iota
	CommentTypeSingleQuote
)

// Comment represents a comment in source code
type Comment struct {
	Type     CommentType
	Text     string
	Range    Range
	Location Location
}

// NewComment creates a new Comment
func NewComment(ctype CommentType, text string) *Comment {
	return &Comment{
		Type: ctype,
		Text: text,
	}
}

// String returns the comment type as string
func (ct CommentType) String() string {
	switch ct {
	case CommentTypeRem:
		return "REM"
	case CommentTypeSingleQuote:
		return "'"
	default:
		return "Unknown"
	}
}

// ParameterModifier represents how a parameter is passed
type ParameterModifier int

const (
	ParameterModifierNone ParameterModifier = iota
	ParameterModifierByRef
	ParameterModifierByVal
)

// Parameter represents a function/sub parameter
type Parameter struct {
	BaseNode
	Modifier    ParameterModifier
	Parentheses bool
	Identifier  *Identifier
}

// NewParameter creates a new Parameter
func NewParameter(id *Identifier, modifier ParameterModifier, parentheses bool) *Parameter {
	if id == nil {
		panic("identifier cannot be nil")
	}
	return &Parameter{
		Modifier:    modifier,
		Parentheses: parentheses,
		Identifier:  id,
	}
}

func (p *Parameter) isStatement() {}

// ConstDeclaration represents a constant declaration
type ConstDeclaration struct {
	BaseNode
	Identifier *Identifier
	Init       Expression
}

// NewConstDeclaration creates a new ConstDeclaration
func NewConstDeclaration(id *Identifier, init Expression) *ConstDeclaration {
	if id == nil {
		panic("identifier cannot be nil")
	}
	if init == nil {
		panic("initialization expression cannot be nil")
	}
	return &ConstDeclaration{
		Identifier: id,
		Init:       init,
	}
}

func (c *ConstDeclaration) isStatement() {}

// ReDimDeclaration represents a ReDim declaration
type ReDimDeclaration struct {
	BaseNode
	Identifier *Identifier
	ArrayDims  []Expression
}

// NewReDimDeclaration creates a new ReDimDeclaration
func NewReDimDeclaration(id *Identifier) *ReDimDeclaration {
	if id == nil {
		panic("identifier cannot be nil")
	}
	return &ReDimDeclaration{
		Identifier: id,
		ArrayDims:  []Expression{},
	}
}

func (r *ReDimDeclaration) isStatement() {}

// VariableDeclarationNode is the base interface for variable declarations
type VariableDeclarationNode interface {
	Node
	isVariableDeclarationNode()
}

// BaseVariableDeclarationNode provides common functionality for variable declarations
type BaseVariableDeclarationNode struct {
	BaseNode
	Identifier     *Identifier
	IsDynamicArray bool
	ArrayDims      []Expression
}

func (v *BaseVariableDeclarationNode) isVariableDeclarationNode() {}
func (v *BaseVariableDeclarationNode) isStatement()               {}

// VariableDeclaration represents a variable declaration
type VariableDeclaration struct {
	BaseVariableDeclarationNode
}

// NewVariableDeclaration creates a new VariableDeclaration
func NewVariableDeclaration(id *Identifier, isDynamicArray bool) *VariableDeclaration {
	if id == nil {
		panic("identifier cannot be nil")
	}
	return &VariableDeclaration{
		BaseVariableDeclarationNode: BaseVariableDeclarationNode{
			Identifier:     id,
			IsDynamicArray: isDynamicArray,
			ArrayDims:      []Expression{},
		},
	}
}

// FieldDeclaration represents a field declaration in a class
type FieldDeclaration struct {
	BaseVariableDeclarationNode
}

// NewFieldDeclaration creates a new FieldDeclaration
func NewFieldDeclaration(id *Identifier, isDynamicArray bool) *FieldDeclaration {
	if id == nil {
		panic("identifier cannot be nil")
	}
	return &FieldDeclaration{
		BaseVariableDeclarationNode: BaseVariableDeclarationNode{
			Identifier:     id,
			IsDynamicArray: isDynamicArray,
			ArrayDims:      []Expression{},
		},
	}
}
