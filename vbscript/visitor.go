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
package vbscript

// Node is the base interface for all AST nodes
type Node interface {
}

// Program represents the root of the AST
type Program interface {
	Node
}

// Parameter represents a function/sub parameter
type Parameter interface {
	Node
}

// ConstDeclaration represents a constant declaration
type ConstDeclaration interface {
	Node
}

// ReDimDeclaration represents a ReDim statement
type ReDimDeclaration interface {
	Node
}

// VariableDeclarationNode is the base interface for variable declarations
type VariableDeclarationNode interface {
	Node
}

// VariableDeclaration represents a variable declaration
type VariableDeclaration interface {
	VariableDeclarationNode
}

// FieldDeclaration represents a field declaration in a class
type FieldDeclaration interface {
	VariableDeclarationNode
}

// Statement is the base interface for all statements
type Statement interface {
	Node
}

// ProcedureDeclaration is the base interface for procedure declarations
type ProcedureDeclaration interface {
	Statement
}

// SubDeclaration represents a Sub declaration
type SubDeclaration interface {
	ProcedureDeclaration
}

// InitializeSubDeclaration represents a Class_Initialize sub
type InitializeSubDeclaration interface {
	SubDeclaration
}

// TerminateSubDeclaration represents a Class_Terminate sub
type TerminateSubDeclaration interface {
	SubDeclaration
}

// FunctionDeclaration represents a Function declaration
type FunctionDeclaration interface {
	ProcedureDeclaration
}

// PropertyDeclaration is the base interface for property declarations
type PropertyDeclaration interface {
	Statement
}

// PropertyGetDeclaration represents a Property Get declaration
type PropertyGetDeclaration interface {
	PropertyDeclaration
}

// PropertySetDeclaration represents a Property Set declaration
type PropertySetDeclaration interface {
	PropertyDeclaration
}

// PropertyLetDeclaration represents a Property Let declaration
type PropertyLetDeclaration interface {
	PropertyDeclaration
}

// Expression is the base interface for all expressions
type Expression interface {
	Node
}

// LiteralExpression is the base interface for literal expressions
type LiteralExpression interface {
	Expression
}

// BooleanLiteral represents a boolean literal (True/False)
type BooleanLiteral interface {
	LiteralExpression
}

// DateLiteral represents a date literal
type DateLiteral interface {
	LiteralExpression
}

// FloatLiteral represents a floating-point literal
type FloatLiteral interface {
	LiteralExpression
}

// NullLiteral represents the Null literal
type NullLiteral interface {
	LiteralExpression
}

// StringLiteral represents a string literal
type StringLiteral interface {
	LiteralExpression
}

// IntegerLiteral represents an integer literal
type IntegerLiteral interface {
	LiteralExpression
}

// NothingLiteral represents the Nothing literal
type NothingLiteral interface {
	LiteralExpression
}

// EmptyLiteral represents the Empty literal
type EmptyLiteral interface {
	LiteralExpression
}

// Identifier represents an identifier expression
type Identifier interface {
	Expression
}

// UnaryExpression represents a unary expression (e.g., Not x, -x)
type UnaryExpression interface {
	Expression
}

// BinaryExpression represents a binary expression (e.g., a + b)
type BinaryExpression interface {
	Expression
}

// IndexOrCallExpression represents array indexing or function call
type IndexOrCallExpression interface {
	Expression
}

// MemberExpression represents member access (e.g., obj.property)
type MemberExpression interface {
	Expression
}

// MissingValueExpression represents a missing value in function call
type MissingValueExpression interface {
	Expression
}

// NewExpression represents a New expression
type NewExpression interface {
	Expression
}

// WithMemberAccessExpression represents member access within a With block
type WithMemberAccessExpression interface {
	Expression
}

// AssignmentStatement represents an assignment statement
type AssignmentStatement interface {
	Statement
}

// CallStatement represents a Call statement
type CallStatement interface {
	Statement
}

// CallSubStatement represents a sub call statement
type CallSubStatement interface {
	Statement
}

// CaseStatement represents a Case statement
type CaseStatement interface {
	Statement
}

// ClassDeclaration represents a Class declaration
type ClassDeclaration interface {
	Statement
}

// ConstsDeclaration represents a Const declaration
type ConstsDeclaration interface {
	Statement
}

// DoStatement represents a Do...Loop statement
type DoStatement interface {
	Statement
}

// ElseIfStatement represents an ElseIf statement
type ElseIfStatement interface {
	Statement
}

// EraseStatement represents an Erase statement
type EraseStatement interface {
	Statement
}

// ExitStatement is the base interface for Exit statements
type ExitStatement interface {
	Statement
}

// ExitDoStatement represents an Exit Do statement
type ExitDoStatement interface {
	ExitStatement
}

// ExitForStatement represents an Exit For statement
type ExitForStatement interface {
	ExitStatement
}

// ExitFunctionStatement represents an Exit Function statement
type ExitFunctionStatement interface {
	ExitStatement
}

// ExitPropertyStatement represents an Exit Property statement
type ExitPropertyStatement interface {
	ExitStatement
}

// ExitSubStatement represents an Exit Sub statement
type ExitSubStatement interface {
	ExitStatement
}

// FieldsDeclaration represents fields declaration
type FieldsDeclaration interface {
	Statement
}

// ForEachStatement represents a For Each...Next statement
type ForEachStatement interface {
	Statement
}

// ForStatement represents a For...Next statement
type ForStatement interface {
	Statement
}

// IfStatement represents an If...Then...Else statement
type IfStatement interface {
	Statement
}

// OnErrorGoTo0Statement represents an On Error GoTo 0 statement
type OnErrorGoTo0Statement interface {
	Statement
}

// OnErrorResumeNextStatement represents an On Error Resume Next statement
type OnErrorResumeNextStatement interface {
	Statement
}

// ReDimStatement represents a ReDim statement
type ReDimStatement interface {
	Statement
}

// SelectStatement represents a Select Case statement
type SelectStatement interface {
	Statement
}

// StatementList represents a list of statements
type StatementList interface {
	Statement
}

// VariablesDeclaration represents variables declaration
type VariablesDeclaration interface {
	Statement
}

// WhileStatement represents a While...Wend statement
type WhileStatement interface {
	Statement
}

// WithStatement represents a With...End With statement
type WithStatement interface {
	Statement
}

// Visitor is a base interface for implementing the Visitor pattern
// It provides Visit methods for all AST node types
type Visitor interface {
	Visit(node Node) interface{}

	VisitProgram(node Program) interface{}
	VisitParameter(node Parameter) interface{}
	VisitConstDeclaration(node ConstDeclaration) interface{}
	VisitReDimDeclaration(node ReDimDeclaration) interface{}

	VisitVariableDeclarationNode(node VariableDeclarationNode) interface{}
	VisitVariableDeclaration(node VariableDeclaration) interface{}
	VisitFieldDeclaration(node FieldDeclaration) interface{}

	VisitStatement(stmt Statement) interface{}

	VisitProcedureDeclaration(stmt ProcedureDeclaration) interface{}
	VisitSubDeclaration(stmt SubDeclaration) interface{}
	VisitInitializeSubDeclaration(stmt InitializeSubDeclaration) interface{}
	VisitTerminateSubDeclaration(stmt TerminateSubDeclaration) interface{}
	VisitFunctionDeclaration(stmt FunctionDeclaration) interface{}

	VisitPropertyDeclaration(stmt PropertyDeclaration) interface{}
	VisitPropertyGetDeclaration(stmt PropertyGetDeclaration) interface{}
	VisitPropertySetDeclaration(stmt PropertySetDeclaration) interface{}
	VisitPropertyLetDeclaration(stmt PropertyLetDeclaration) interface{}

	VisitExpression(expr Expression) interface{}

	VisitLiteralExpression(expr LiteralExpression) interface{}
	VisitBooleanLiteral(expr BooleanLiteral) interface{}
	VisitDateLiteral(expr DateLiteral) interface{}
	VisitFloatLiteral(expr FloatLiteral) interface{}
	VisitNullLiteral(expr NullLiteral) interface{}
	VisitStringLiteral(expr StringLiteral) interface{}
	VisitIntegerLiteral(expr IntegerLiteral) interface{}
	VisitNothingLiteral(expr NothingLiteral) interface{}
	VisitEmptyLiteral(expr EmptyLiteral) interface{}

	VisitIdentifier(expr Identifier) interface{}
	VisitUnaryExpression(expr UnaryExpression) interface{}
	VisitBinaryExpression(expr BinaryExpression) interface{}
	VisitIndexOrCallExpression(expr IndexOrCallExpression) interface{}
	VisitMemberExpression(expr MemberExpression) interface{}
	VisitMissingValueExpression(expr MissingValueExpression) interface{}
	VisitNewExpression(expr NewExpression) interface{}
	VisitWithMemberAccessExpression(expr WithMemberAccessExpression) interface{}

	VisitAssignmentStatement(stmt AssignmentStatement) interface{}
	VisitCallStatement(stmt CallStatement) interface{}
	VisitCallSubStatement(stmt CallSubStatement) interface{}
	VisitCaseStatement(stmt CaseStatement) interface{}
	VisitClassDeclaration(stmt ClassDeclaration) interface{}
	VisitConstsDeclaration(stmt ConstsDeclaration) interface{}
	VisitDoStatement(stmt DoStatement) interface{}
	VisitElseIfStatement(stmt ElseIfStatement) interface{}
	VisitEraseStatement(stmt EraseStatement) interface{}

	VisitExitStatement(stmt ExitStatement) interface{}
	VisitExitDoStatement(stmt ExitDoStatement) interface{}
	VisitExitForStatement(stmt ExitForStatement) interface{}
	VisitExitFunctionStatement(stmt ExitFunctionStatement) interface{}
	VisitExitPropertyStatement(stmt ExitPropertyStatement) interface{}
	VisitExitSubStatement(stmt ExitSubStatement) interface{}

	VisitFieldsDeclaration(stmt FieldsDeclaration) interface{}
	VisitForEachStatement(stmt ForEachStatement) interface{}
	VisitForStatement(stmt ForStatement) interface{}
	VisitIfStatement(stmt IfStatement) interface{}
	VisitOnErrorGoTo0Statement(stmt OnErrorGoTo0Statement) interface{}
	VisitOnErrorResumeNextStatement(stmt OnErrorResumeNextStatement) interface{}

	VisitReDimStatement(stmt ReDimStatement) interface{}
	VisitSelectStatement(stmt SelectStatement) interface{}
	VisitStatementList(stmt StatementList) interface{}

	VisitVariablesDeclaration(stmt VariablesDeclaration) interface{}
	VisitWhileStatement(stmt WhileStatement) interface{}
	VisitWithStatement(stmt WithStatement) interface{}
}
