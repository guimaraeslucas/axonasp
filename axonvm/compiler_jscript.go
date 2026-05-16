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
package axonvm

import (
	"encoding/binary"
	"fmt"
	"math"
	"math/big"
	"regexp"
	"strconv"
	"strings"

	"g3pix.com.br/axonasp/jscript"
	jsast "g3pix.com.br/axonasp/jscript/ast"
	jsparser "g3pix.com.br/axonasp/jscript/parser"
	jstoken "g3pix.com.br/axonasp/jscript/token"
	jsunistring "g3pix.com.br/axonasp/jscript/unistring"
)

var jscriptCallAssignmentPattern = regexp.MustCompile(`([A-Za-z_][A-Za-z0-9_]*)\s*\(([^\)]*)\)\s*=\s+([^;\r\n]+);`)

const jsRestParamTemplatePrefix = "__js_rest__:"

// compileJScriptBlock parses one JScript source block and emits isolated OpJS bytecode.
func (c *Compiler) compileJScriptBlock(source string) {
	// Classic ASP JScript commonly uses indexed default-property assignment syntax
	// like Session("key") = value; normalize it into Session("key", value);
	// so the GoJa parser accepts it and dispatchNativeCall(member="") can execute it.
	source = jscriptCallAssignmentPattern.ReplaceAllString(source, `$1($2, $3);`)

	mode := jsparser.Mode(0)
	if c.isJSModule {
		mode |= jsparser.ModeModule
	}

	program, err := jsparser.ParseFile(nil, c.sourceName, source, mode)
	if err != nil {
		panic(c.newJScriptCompileErrorFromParse(err, "jscript parse error"))
	}

	prevLocalEnabled := c.jsLocalEnabled
	prevLocalSlotCount := c.jsLocalSlotCount
	prevLocalScopeStack := c.jsLocalScopeStack
	c.jsLocalEnabled = !c.isJSModule && !jsProgramContainsNestedFunction(program.Body)
	c.jsLocalSlotCount = 0
	c.jsLocalScopeStack = make([]jsLocalScope, 0, 8)
	if c.jsLocalEnabled {
		c.jsPushLocalScope(false)
	}
	rootFrameEnterPos := -1
	if c.jsLocalEnabled {
		rootFrameEnterPos = c.emit(OpJSRootFrameEnter, 0)
	}
	defer func() {
		c.jsLocalEnabled = prevLocalEnabled
		c.jsLocalSlotCount = prevLocalSlotCount
		c.jsLocalScopeStack = prevLocalScopeStack
	}()

	// Detect "use strict" directive at the beginning of the program
	hasStrictMode, _ := c.detectUseStrictDirective(program.Body)
	if hasStrictMode {
		c.emit(OpJSStrictModeEnter)
		prevStrictMode := c.jsStrictMode
		c.jsStrictMode = true
		defer func() { c.jsStrictMode = prevStrictMode }()
	}

	// Top-level script/module scope has lexical semantics for let/const/class.
	// Register these names in a dedicated block scope so const reassignment and
	// TDZ checks behave like ES6 outside nested blocks as well.
	topLetNames, topConstNames := jsGetBlockLexicalNames(program.Body)
	hasTopLexical := len(topLetNames) > 0 || len(topConstNames) > 0
	if hasTopLexical {
		c.emit(OpJSBlockScopeEnter)
		for _, name := range topLetNames {
			if c.jsLocalEnabled {
				c.jsAddLocalBarrier(name)
			}
			c.emit(OpJSTDZRegisterLet, c.addConstant(NewString(name)))
		}
		for _, name := range topConstNames {
			if c.jsLocalEnabled {
				c.jsAddLocalBarrier(name)
			}
			c.emit(OpJSTDZRegisterConst, c.addConstant(NewString(name)))
		}
	}

	c.compileJScriptScopedStatements(program.Body)

	if hasTopLexical {
		c.emit(OpJSBlockScopeExit)
	}

	if c.jsLocalEnabled {
		if rootFrameEnterPos >= 0 {
			binary.BigEndian.PutUint16(c.bytecode[rootFrameEnterPos+1:], uint16(c.jsLocalSlotCount))
		}
		c.emit(OpJSRootFrameLeave, c.jsLocalSlotCount)
	}

	if hasStrictMode {
		c.emit(OpJSStrictModeExit)
	}
}

// compileJScriptEvalSnippet parses one JScript eval source and emits OpJS bytecode
// that leaves the completion value on the stack and terminates with OpHalt.
func (c *Compiler) compileJScriptEvalSnippet(source string) {
	source = jscriptCallAssignmentPattern.ReplaceAllString(source, `$1($2, $3);`)

	program, err := jsparser.ParseFile(nil, c.sourceName, source, 0)
	if err != nil {
		panic(c.newJScriptCompileErrorFromParse(err, "jscript eval parse error"))
	}

	if len(program.Body) == 0 {
		c.emit(OpJSLoadUndefined)
		c.emit(OpHalt)
		return
	}

	// Detect "use strict" directive at the beginning
	hasStrictMode, _ := c.detectUseStrictDirective(program.Body)
	if hasStrictMode {
		c.emit(OpJSStrictModeEnter)
		prevStrictMode := c.jsStrictMode
		c.jsStrictMode = true
		defer func() { c.jsStrictMode = prevStrictMode }()
	}

	// Eval code also executes in a lexical program scope for let/const/class.
	topLetNames, topConstNames := jsGetBlockLexicalNames(program.Body)
	hasTopLexical := len(topLetNames) > 0 || len(topConstNames) > 0
	if hasTopLexical {
		c.emit(OpJSBlockScopeEnter)
		for _, name := range topLetNames {
			if c.jsLocalEnabled {
				c.jsAddLocalBarrier(name)
			}
			c.emit(OpJSTDZRegisterLet, c.addConstant(NewString(name)))
		}
		for _, name := range topConstNames {
			if c.jsLocalEnabled {
				c.jsAddLocalBarrier(name)
			}
			c.emit(OpJSTDZRegisterConst, c.addConstant(NewString(name)))
		}
	}

	lastIdx := len(program.Body) - 1
	if lastIdx > 0 {
		c.compileJScriptScopedStatements(program.Body[:lastIdx])
	}

	last := program.Body[lastIdx]
	if exprStmt, ok := last.(*jsast.ExpressionStatement); ok {
		c.compileJScriptExpression(exprStmt.Expression)
	} else {
		c.compileJScriptStatement(last)
		c.emit(OpJSLoadUndefined)
	}

	if hasTopLexical {
		c.emit(OpJSBlockScopeExit)
	}

	if hasStrictMode {
		c.emit(OpJSStrictModeExit)
	}

	c.emit(OpHalt)
}

// newJScriptCompileErrorFromParse converts parser failures into a JScript syntax error.
func (c *Compiler) newJScriptCompileErrorFromParse(parseErr error, detailPrefix string) *jscript.JSSyntaxError {
	detail := strings.TrimSpace(parseErr.Error())
	line, col := 0, 0
	if parserErr, ok := parseErr.(*jsparser.Error); ok {
		line = parserErr.Position.Line
		col = parserErr.Position.Column
		detail = strings.TrimSpace(parserErr.Message)
	} else if parserList, ok := parseErr.(jsparser.ErrorList); ok && len(parserList) > 0 {
		line = parserList[0].Position.Line
		col = parserList[0].Position.Column
		detail = strings.TrimSpace(parserList[0].Message)
	}

	jsErr := jscript.NewJSSyntaxError(jscript.SyntaxError, line, col)
	if detail == "" {
		jsErr.WithASPDescription(detailPrefix)
	} else {
		jsErr.WithASPDescription(detailPrefix + ": " + detail)
	}
	if c.sourceName != "" {
		jsErr.WithFile(c.sourceName)
	}
	return jsErr
}

// detectUseStrictDirective checks if the first statement(s) contain a "use strict" directive.
// A directive is a StringLiteral ExpressionStatement before any other statement type.
// Returns true if strict mode is enabled, and the number of directive statements to skip.
func (c *Compiler) detectUseStrictDirective(statements []jsast.Statement) (hasStrictMode bool, directiveCount int) {
	directiveCount = 0
	for i, stmt := range statements {
		exprStmt, ok := stmt.(*jsast.ExpressionStatement)
		if !ok {
			break
		}
		strLit, ok := exprStmt.Expression.(*jsast.StringLiteral)
		if !ok {
			break
		}
		if strings.EqualFold(strLit.Value.String(), "use strict") {
			hasStrictMode = true
			directiveCount = i + 1
		} else {
			// Other directive strings are allowed but ignored
			directiveCount = i + 1
		}
	}
	return
}

// pushJSLoopContext adds a new loop context to the stack.
func (c *Compiler) pushJSLoopContext() *jsLoopContext {
	ctx := &jsLoopContext{
		continueTargets:     make([]int, 0),
		loopStart:           len(c.bytecode),
		forIterDepthAtStart: len(c.jsForIterScopes),
	}
	if c.jsLoopContexts == nil {
		c.jsLoopContexts = make([]*jsLoopContext, 0)
	}
	c.jsLoopContexts = append(c.jsLoopContexts, ctx)
	return ctx
}

// popJSLoopContext removes the current loop context from the stack.
func (c *Compiler) popJSLoopContext() *jsLoopContext {
	if len(c.jsLoopContexts) == 0 {
		return nil
	}
	ctx := c.jsLoopContexts[len(c.jsLoopContexts)-1]
	c.jsLoopContexts = c.jsLoopContexts[:len(c.jsLoopContexts)-1]
	return ctx
}

// currentJSLoopContext returns the current loop context or nil if not in a loop.
func (c *Compiler) currentJSLoopContext() *jsLoopContext {
	if len(c.jsLoopContexts) == 0 {
		return nil
	}
	return c.jsLoopContexts[len(c.jsLoopContexts)-1]
}

func (c *Compiler) pushJSBreakContext() *jsBreakContext {
	ctx := &jsBreakContext{
		breakTargets:        make([]int, 0),
		forIterDepthAtStart: len(c.jsForIterScopes),
	}
	if c.jsBreakContexts == nil {
		c.jsBreakContexts = make([]*jsBreakContext, 0)
	}
	c.jsBreakContexts = append(c.jsBreakContexts, ctx)
	return ctx
}

func (c *Compiler) popJSBreakContext() *jsBreakContext {
	if len(c.jsBreakContexts) == 0 {
		return nil
	}
	ctx := c.jsBreakContexts[len(c.jsBreakContexts)-1]
	c.jsBreakContexts = c.jsBreakContexts[:len(c.jsBreakContexts)-1]
	return ctx
}

func (c *Compiler) currentJSBreakContext() *jsBreakContext {
	if len(c.jsBreakContexts) == 0 {
		return nil
	}
	return c.jsBreakContexts[len(c.jsBreakContexts)-1]
}

func (c *Compiler) jsPushLocalScope(isFunction bool) {
	if !c.jsLocalEnabled {
		return
	}
	scope := jsLocalScope{
		entries:    make(map[string]int, 8),
		types:      make(map[string]jsType, 8),
		isFunction: isFunction,
	}
	c.jsLocalScopeStack = append(c.jsLocalScopeStack, scope)
}

func (c *Compiler) jsSetLocalType(name string, t jsType) {
	if !c.jsLocalEnabled {
		return
	}
	for i := len(c.jsLocalScopeStack) - 1; i >= 0; i-- {
		scope := c.jsLocalScopeStack[i]
		if _, exists := scope.entries[name]; exists {
			scope.types[name] = t
			return
		}
	}
}

func (c *Compiler) jsGetLocalType(name string) jsType {
	if !c.jsLocalEnabled {
		return jsTypeUnknown
	}
	for i := len(c.jsLocalScopeStack) - 1; i >= 0; i-- {
		scope := c.jsLocalScopeStack[i]
		if t, exists := scope.types[name]; exists {
			return t
		}
	}
	return jsTypeUnknown
}

func (c *Compiler) jsPopLocalScope() {
	if !c.jsLocalEnabled || len(c.jsLocalScopeStack) == 0 {
		return
	}
	c.jsLocalScopeStack = c.jsLocalScopeStack[:len(c.jsLocalScopeStack)-1]
}

func (c *Compiler) jsAddLocalBarrier(name string) {
	if !c.jsLocalEnabled || len(c.jsLocalScopeStack) == 0 {
		return
	}
	c.jsLocalScopeStack[len(c.jsLocalScopeStack)-1].entries[name] = -1
}

func (c *Compiler) jsDeclareFunctionLocal(name string) int {
	if !c.jsLocalEnabled {
		return -1
	}
	for i := len(c.jsLocalScopeStack) - 1; i >= 0; i-- {
		scope := &c.jsLocalScopeStack[i]
		if !scope.isFunction {
			continue
		}
		if slot, exists := scope.entries[name]; exists && slot >= 0 {
			return slot
		}
		slot := c.jsLocalSlotCount
		c.jsLocalSlotCount++
		scope.entries[name] = slot
		return slot
	}
	return -1
}

func (c *Compiler) jsDeclareCurrentLocal(name string) int {
	if !c.jsLocalEnabled || len(c.jsLocalScopeStack) == 0 {
		return -1
	}
	scope := &c.jsLocalScopeStack[len(c.jsLocalScopeStack)-1]
	if slot, exists := scope.entries[name]; exists && slot >= 0 {
		return slot
	}
	if c.jsHasFunctionLocalScope() {
		slot := c.jsLocalSlotCount
		c.jsLocalSlotCount++
		scope.entries[name] = slot
		return slot
	}

	// Root JScript bytecode has no dedicated call frame, so reserve one stable
	// offset relative to vm.fp. Root slot memory is provisioned by OpJSRootFrameEnter.
	slot := c.jsLocalSlotCount
	c.jsLocalSlotCount++
	scope.entries[name] = slot
	return slot
}

func (c *Compiler) jsHasFunctionLocalScope() bool {
	for i := len(c.jsLocalScopeStack) - 1; i >= 0; i-- {
		if c.jsLocalScopeStack[i].isFunction {
			return true
		}
	}
	return false
}

func (c *Compiler) jsResolveLocalSlot(name string) (int, bool) {
	if !c.jsLocalEnabled {
		return 0, false
	}
	for i := len(c.jsLocalScopeStack) - 1; i >= 0; i-- {
		scope := c.jsLocalScopeStack[i]
		if slot, exists := scope.entries[name]; exists {
			if slot < 0 {
				return 0, false
			}
			return slot, true
		}
	}
	return 0, false
}

func (c *Compiler) jsInferredType(expr jsast.Expression) jsType {
	switch node := expr.(type) {
	case *jsast.NumberLiteral:
		if _, ok := jsNumericLiteralInt64(node); ok {
			return jsTypeInteger
		}
	case *jsast.Identifier:
		return c.jsGetLocalType(node.Name.String())
	case *jsast.BinaryExpression:
		switch node.Operator {
		case jstoken.OR, jstoken.AND, jstoken.EXCLUSIVE_OR, jstoken.SHIFT_LEFT, jstoken.SHIFT_RIGHT, jstoken.UNSIGNED_SHIFT_RIGHT:
			return jsTypeInteger
		case jstoken.PLUS, jstoken.MINUS, jstoken.MULTIPLY:
			if c.jsInferredType(node.Left) == jsTypeInteger && c.jsInferredType(node.Right) == jsTypeInteger {
				return jsTypeInteger
			}
		}
	case *jsast.UnaryExpression:
		if node.Operator == jstoken.INCREMENT || node.Operator == jstoken.DECREMENT {
			if id, ok := node.Operand.(*jsast.Identifier); ok {
				return c.jsGetLocalType(id.Name.String())
			}
		}
	case *jsast.CallExpression:
		if callee, ok := node.Callee.(*jsast.DotExpression); ok {
			if id, ok := callee.Left.(*jsast.Identifier); ok && id.Name.String() == "Math" {
				method := callee.Identifier.Name.String()
				switch method {
				case "abs", "floor", "ceil", "round":
					return jsTypeInteger
				}
			}
		}
	}
	return jsTypeUnknown
}

func (c *Compiler) compileJScriptStatement(stmt jsast.Statement) {
	switch node := stmt.(type) {
	case *jsast.ExpressionStatement:
		c.compileJScriptExpression(node.Expression)
		c.emit(OpJSPop)
	case *jsast.VariableStatement:
		for _, binding := range node.List {
			if binding.Initializer != nil {
				c.compileJScriptExpression(binding.Initializer)
				t := jsTypeUnknown
				if c.jsInferredType(binding.Initializer) == jsTypeInteger {
					t = jsTypeInteger
				}
				c.compileJScriptDestructuring(binding.Target, false, false, true)
				if t == jsTypeInteger {
					if id, ok := binding.Target.(*jsast.Identifier); ok {
						c.jsSetLocalType(id.Name.String(), jsTypeInteger)
					}
				}
			} else {
				// var x; -> declare x
				if id, ok := binding.Target.(*jsast.Identifier); ok {
					if slot, hasLocal := c.jsResolveLocalSlot(id.Name.String()); hasLocal {
						_ = slot
						continue
					}
					if c.jsLocalEnabled {
						if slot := c.jsDeclareFunctionLocal(id.Name.String()); slot >= 0 {
							continue
						}
					}
					nameIdx := c.addConstant(NewString(id.Name.String()))
					c.emit(OpJSDeclareName, nameIdx)
				}
			}
		}
	case *jsast.LexicalDeclaration:
		// Handle ES6 let/const declarations with block scoping
		c.compileJScriptLexicalDeclaration(node)
	case *jsast.UsingDeclaration:
		// Fallback path for top-level or non-block contexts.
		c.compileJScriptUsingDeclaration(node, nil)
	case *jsast.FunctionDeclaration:
		if node.Function == nil {
			return
		}
		name := ""
		if node.Function.Name != nil {
			name = node.Function.Name.Name.String()
		}
		if name == "" {
			return
		}
		nameIdx := c.addConstant(NewString(name))
		c.emit(OpJSDeclareName, nameIdx)
		c.compileJScriptFunctionLiteral(node.Function, name, false)
		c.emit(OpJSSetName, nameIdx)
	case *jsast.ClassDeclaration:
		c.compileJScriptClassDeclaration(node)
	case *jsast.ImportDeclaration:
		c.compileJScriptImportDeclaration(node)
	case *jsast.ExportDeclaration:
		c.compileJScriptExportDeclaration(node)
	case *jsast.ReturnStatement:
		c.emitJSLeaveWithScopes(c.withDepth)
		if c.jsTryDepth == 0 && c.compileJScriptTailReturn(node.Argument) {
			return
		}
		if node.Argument != nil {
			c.compileJScriptExpression(node.Argument)
		} else {
			c.emit(OpJSLoadUndefined)
		}
		c.emit(OpJSReturn)
	case *jsast.ThrowStatement:
		c.emitJSLeaveWithScopes(c.withDepth)
		if node.Argument != nil {
			c.compileJScriptExpression(node.Argument)
		} else {
			c.emit(OpJSLoadUndefined)
		}
		c.emit(OpJSThrow)
	case *jsast.BlockStatement:
		// Check if block contains let/const declarations requiring a block scope
		letNames, constNames := jsGetBlockLexicalNames(node.List)
		hasLexical := len(letNames) > 0 || len(constNames) > 0
		if c.jsLocalEnabled {
			c.jsPushLocalScope(false)
			for _, name := range letNames {
				c.jsAddLocalBarrier(name)
			}
			for _, name := range constNames {
				c.jsAddLocalBarrier(name)
			}
		}
		if hasLexical {
			c.emit(OpJSBlockScopeEnter)
			for _, name := range letNames {
				c.emit(OpJSTDZRegisterLet, c.addConstant(NewString(name)))
			}
			for _, name := range constNames {
				c.emit(OpJSTDZRegisterConst, c.addConstant(NewString(name)))
			}
		}
		c.compileJScriptScopedStatements(node.List)
		if hasLexical {
			c.emit(OpJSBlockScopeExit)
		}
		if c.jsLocalEnabled {
			c.jsPopLocalScope()
		}
	case *jsast.TryStatement:
		c.jsTryDepth++
		tryPos := c.emit(OpJSTryEnter, 0)
		c.compileJScriptStatement(node.Body)
		c.emit(OpJSTryLeave)
		jumpEnd := c.emitJSJump(OpJSJump)
		c.patchJSJumpTo(tryPos+1, len(c.bytecode))
		if node.Catch != nil {
			if id, ok := node.Catch.Parameter.(*jsast.Identifier); ok {
				if c.jsLocalEnabled {
					c.jsPushLocalScope(false)
					c.jsAddLocalBarrier(id.Name.String())
				}
				nameIdx := c.addConstant(NewString(id.Name.String()))
				c.emit(OpJSDeclareName, nameIdx)
				c.emit(OpJSLoadCatchError)
				c.emit(OpJSSetName, nameIdx)
				c.compileJScriptStatement(node.Catch.Body)
				if c.jsLocalEnabled {
					c.jsPopLocalScope()
				}
			} else {
				c.compileJScriptStatement(node.Catch.Body)
			}
		}
		if node.Finally != nil {
			c.compileJScriptStatement(node.Finally)
		}
		c.jsTryDepth--
		c.patchJSJump(jumpEnd)
	case *jsast.IfStatement:
		c.compileJScriptExpression(node.Test)
		jumpFalse := c.emitJSJump(OpJSJumpIfFalse)
		c.compileJScriptStatement(node.Consequent)
		jumpEnd := c.emitJSJump(OpJSJump)
		c.patchJSJump(jumpFalse)
		if node.Alternate != nil {
			c.compileJScriptStatement(node.Alternate)
		}
		c.patchJSJump(jumpEnd)
	case *jsast.WhileStatement:
		c.compileJScriptWhileStatement(node)
	case *jsast.DoWhileStatement:
		c.compileJScriptDoWhileStatement(node)
	case *jsast.ForStatement:
		c.compileJScriptForStatement(node)
	case *jsast.ForInStatement:
		c.compileJScriptForInStatement(node)
	case *jsast.ForOfStatement:
		c.compileJScriptForOfStatement(node)
	case *jsast.WithStatement:
		c.compileJScriptWithStatement(node)
	case *jsast.BranchStatement:
		// BranchStatement handles both break and continue
		switch node.Token {
		case jstoken.BREAK:
			breakCtx := c.currentJSBreakContext()
			if breakCtx != nil {
				c.emitJSLeaveWithScopes(c.withDepth)
				// Exit any active per-iteration scopes within this break context
				c.emitJSLeaveForIterScopes(breakCtx.forIterDepthAtStart)
				jumpPos := c.emitJSJump(OpJSBreak)
				breakCtx.breakTargets = append(breakCtx.breakTargets, jumpPos)
			}
		case jstoken.CONTINUE:
			loopCtx := c.currentJSLoopContext()
			if loopCtx != nil {
				c.emitJSLeaveWithScopes(c.withDepth)
				// Exit any active per-iteration scopes within this loop context
				c.emitJSLeaveForIterScopes(loopCtx.forIterDepthAtStart)
				jumpPos := c.emitJSJump(OpJSContinue)
				loopCtx.continueTargets = append(loopCtx.continueTargets, jumpPos)
			}
		}
	case *jsast.SwitchStatement:
		c.compileJScriptSwitchStatement(node)
	}
}

type jsUsingBinding struct {
	name     string
	symbolID int64
}

func jsStatementHasUsingDeclaration(stmts []jsast.Statement) bool {
	for i := range stmts {
		if _, ok := stmts[i].(*jsast.UsingDeclaration); ok {
			return true
		}
	}
	return false
}

func (c *Compiler) emitJScriptDisposeBindings(bindings []jsUsingBinding) {
	for i := len(bindings) - 1; i >= 0; i-- {
		nameIdx := c.addConstant(NewString(bindings[i].name))
		symbolKey := jsSymbolPropertyPrefix + strconv.FormatInt(bindings[i].symbolID, 10)
		symbolKeyIdx := c.addConstant(NewString(symbolKey))
		c.emit(OpJSGetName, nameIdx)
		c.emit(OpConstant, symbolKeyIdx)
		c.emit(OpJSCallComputedMember, 0)
		c.emit(OpJSPop)
	}
}

func (c *Compiler) compileJScriptScopedStatements(stmts []jsast.Statement) {
	if !jsStatementHasUsingDeclaration(stmts) {
		for i := range stmts {
			c.compileJScriptStatement(stmts[i])
		}
		return
	}

	usingBindings := make([]jsUsingBinding, 0, 4)
	c.jsTryDepth++
	tryPos := c.emit(OpJSTryEnter, 0)
	for i := range stmts {
		if usingDecl, ok := stmts[i].(*jsast.UsingDeclaration); ok {
			c.compileJScriptUsingDeclaration(usingDecl, &usingBindings)
			continue
		}
		c.compileJScriptStatement(stmts[i])
	}
	c.emit(OpJSTryLeave)
	c.emitJScriptDisposeBindings(usingBindings)
	jumpEnd := c.emitJSJump(OpJSJump)

	catchStart := len(c.bytecode)
	c.patchJSJumpTo(tryPos+1, catchStart)

	errName := fmt.Sprintf("__js_using_err_%d", c.tempCounter)
	c.tempCounter++
	errNameIdx := c.addConstant(NewString(errName))
	c.emit(OpJSDeclareName, errNameIdx)
	c.emit(OpJSLoadCatchError)
	c.emit(OpJSSetName, errNameIdx)
	c.emitJScriptDisposeBindings(usingBindings)
	c.emit(OpJSGetName, errNameIdx)
	c.emit(OpJSThrow)

	c.jsTryDepth--
	c.patchJSJump(jumpEnd)
}

func (c *Compiler) compileJScriptUsingDeclaration(node *jsast.UsingDeclaration, bindings *[]jsUsingBinding) {
	symbolID := int64(jsWellKnownSymbolDispose)
	if node.IsAsync {
		symbolID = int64(jsWellKnownSymbolAsyncDispose)
	}

	for _, binding := range node.List {
		if binding.Initializer != nil {
			c.compileJScriptExpression(binding.Initializer)
			t := jsTypeUnknown
			if c.jsInferredType(binding.Initializer) == jsTypeInteger {
				t = jsTypeInteger
			}
			c.compileJScriptDestructuring(binding.Target, false, true, false)
			if t == jsTypeInteger {
				if id, ok := binding.Target.(*jsast.Identifier); ok {
					c.jsSetLocalType(id.Name.String(), jsTypeInteger)
				}
			}
		} else {
			if id, ok := binding.Target.(*jsast.Identifier); ok {
				if c.jsLocalEnabled {
					c.jsAddLocalBarrier(id.Name.String())
				}
				nameIdx := c.addConstant(NewString(id.Name.String()))
				c.emit(OpJSLetDeclare, nameIdx)
			}
		}

		if bindings != nil {
			names := make([]string, 0, 2)
			jsExtractBindingNames(binding.Target, &names)
			for j := range names {
				*bindings = append(*bindings, jsUsingBinding{name: names[j], symbolID: symbolID})
			}
		}
	}
}

func (c *Compiler) compileJScriptImportDeclaration(node *jsast.ImportDeclaration) {
	if node == nil || node.Source == nil {
		return
	}
	moduleIdx := c.addConstant(NewString(node.Source.Value.String()))
	c.bytecode = append(c.bytecode, byte(OpJSImport))
	c.bytecode = append(c.bytecode, byte(moduleIdx>>8), byte(moduleIdx&0xFF))
	specCount := len(node.Specifiers)
	c.bytecode = append(c.bytecode, byte(specCount>>8), byte(specCount&0xFF))
	for i := 0; i < specCount; i++ {
		importedName := ""
		localName := ""
		if node.Specifiers[i].IsDefault {
			importedName = "default"
		} else if node.Specifiers[i].IsNamespace {
			importedName = "*"
		} else if node.Specifiers[i].Imported != nil {
			importedName = node.Specifiers[i].Imported.Name.String()
		}
		if node.Specifiers[i].Local != nil {
			localName = node.Specifiers[i].Local.Name.String()
		}
		importedIdx := c.addConstant(NewString(importedName))
		localIdx := c.addConstant(NewString(localName))
		c.bytecode = append(c.bytecode, byte(importedIdx>>8), byte(importedIdx&0xFF))
		c.bytecode = append(c.bytecode, byte(localIdx>>8), byte(localIdx&0xFF))
	}
}

func jsCollectDeclarationBindingNames(stmt jsast.Statement) []string {
	if stmt == nil {
		return nil
	}
	switch decl := stmt.(type) {
	case *jsast.VariableStatement:
		names := make([]string, 0, len(decl.List))
		for i := 0; i < len(decl.List); i++ {
			if id, ok := decl.List[i].Target.(*jsast.Identifier); ok {
				names = append(names, id.Name.String())
			}
		}
		return names
	case *jsast.LexicalDeclaration:
		names := make([]string, 0, len(decl.List))
		for i := 0; i < len(decl.List); i++ {
			if id, ok := decl.List[i].Target.(*jsast.Identifier); ok {
				names = append(names, id.Name.String())
			}
		}
		return names
	case *jsast.FunctionDeclaration:
		if decl.Function != nil && decl.Function.Name != nil {
			return []string{decl.Function.Name.Name.String()}
		}
	case *jsast.ClassDeclaration:
		if decl.Class != nil && decl.Class.Name != nil {
			return []string{decl.Class.Name.Name.String()}
		}
	}
	return nil
}

func (c *Compiler) emitJScriptExport(localName string, exportName string) {
	localIdx := c.addConstant(NewString(localName))
	exportIdx := c.addConstant(NewString(exportName))
	c.emit(OpJSExport, localIdx, exportIdx)
}

func (c *Compiler) compileJScriptExportDeclaration(node *jsast.ExportDeclaration) {
	if node == nil {
		return
	}

	if node.IsDefault {
		if node.Declaration != nil {
			names := jsCollectDeclarationBindingNames(node.Declaration)
			if len(names) > 0 {
				c.compileJScriptStatement(node.Declaration)
				c.emitJScriptExport(names[0], "default")
			} else {
				localAlias := fmt.Sprintf("__js_export_default_tmp__%d", c.tempCounter)
				c.tempCounter++
				c.emit(OpJSDeclareName, c.addConstant(NewString(localAlias)))
				if fn, ok := node.Declaration.(*jsast.FunctionDeclaration); ok {
					c.compileJScriptFunctionLiteral(fn.Function, "", false)
				} else if cl, ok := node.Declaration.(*jsast.ClassDeclaration); ok {
					c.compileJScriptClassLiteral(cl.Class)
				} else if ex, ok := node.Declaration.(*jsast.ExpressionStatement); ok {
					c.compileJScriptExpression(ex.Expression)
				} else {
					// Fallback
					c.compileJScriptStatement(node.Declaration)
					c.emit(OpJSLoadUndefined) // Should not happen for valid ES6
				}
				c.emit(OpJSSetName, c.addConstant(NewString(localAlias)))
				c.emitJScriptExport(localAlias, "default")
			}
		}
		return
	}

	if node.IsAll && node.Source != nil {
		if len(node.Specifiers) == 1 {
			exportName := node.Specifiers[0].Exported.Name.String()
			localAlias := fmt.Sprintf("__js_export_ns_tmp__%d", c.tempCounter)
			c.tempCounter++
			imp := &jsast.ImportDeclaration{
				Source: node.Source,
				Specifiers: []jsast.JSImportSpecifier{
					{Local: &jsast.Identifier{Name: jsunistring.String(localAlias)}, IsNamespace: true},
				},
			}
			c.compileJScriptImportDeclaration(imp)
			c.emitJScriptExport(localAlias, exportName)
		} else {
			moduleIdx := c.addConstant(NewString(node.Source.Value.String()))
			c.emit(OpJSExportAll, moduleIdx)
		}
		return
	}

	if node.Declaration != nil {
		c.compileJScriptStatement(node.Declaration)
		names := jsCollectDeclarationBindingNames(node.Declaration)
		for i := 0; i < len(names); i++ {
			c.emitJScriptExport(names[i], names[i])
		}
		return
	}

	if len(node.Specifiers) == 0 {
		return
	}

	if node.Source != nil {
		for i := 0; i < len(node.Specifiers); i++ {
			importedName := ""
			localAlias := ""
			exportName := ""
			if node.Specifiers[i].Local != nil {
				importedName = node.Specifiers[i].Local.Name.String()
			}
			if node.Specifiers[i].Exported != nil {
				exportName = node.Specifiers[i].Exported.Name.String()
			}
			localAlias = fmt.Sprintf("__js_export_tmp__%d", c.tempCounter)
			c.tempCounter++
			imp := &jsast.ImportDeclaration{
				Source: node.Source,
				Specifiers: []jsast.JSImportSpecifier{
					{Imported: &jsast.Identifier{Name: jsunistring.String(importedName)}, Local: &jsast.Identifier{Name: jsunistring.String(localAlias)}},
				},
			}
			c.compileJScriptImportDeclaration(imp)
			c.emitJScriptExport(localAlias, exportName)
		}
		return
	}

	for i := 0; i < len(node.Specifiers); i++ {
		localName := ""
		exportName := ""
		if node.Specifiers[i].Local != nil {
			localName = node.Specifiers[i].Local.Name.String()
		}
		if node.Specifiers[i].Exported != nil {
			exportName = node.Specifiers[i].Exported.Name.String()
		}
		c.emitJScriptExport(localName, exportName)
	}
}

// emitJSLeaveWithScopes emits OpWithLeave for each active JScript with-scope.
func (c *Compiler) emitJSLeaveWithScopes(count int) {
	for i := 0; i < count; i++ {
		c.emit(OpWithLeave)
	}
}

// emitJSLeaveForIterScopes emits OpJSForIterExit for all active per-iteration scopes
// above targetDepth (the depth at the start of the enclosing loop/break context).
func (c *Compiler) emitJSLeaveForIterScopes(targetDepth int) {
	for i := len(c.jsForIterScopes) - 1; i >= targetDepth; i-- {
		if c.jsForIterScopes[i].fast {
			c.emitForIterExitFast()
		} else {
			c.emitForIterExit(c.jsForIterScopes[i].nameIdxs)
		}
	}
}

// emitForIterEnter emits OpJSForIterEnter with the given variable name constant indices.
func (c *Compiler) emitForIterEnter(nameIdxs []int) {
	c.bytecode = append(c.bytecode, byte(OpJSForIterEnter))
	numVars := len(nameIdxs)
	c.bytecode = append(c.bytecode, byte(numVars>>8), byte(numVars&0xFF))
	for _, idx := range nameIdxs {
		c.bytecode = append(c.bytecode, byte(idx>>8), byte(idx&0xFF))
	}
}

// emitForIterExit emits OpJSForIterExit with the given variable name constant indices.
func (c *Compiler) emitForIterExit(nameIdxs []int) {
	c.bytecode = append(c.bytecode, byte(OpJSForIterExit))
	numVars := len(nameIdxs)
	c.bytecode = append(c.bytecode, byte(numVars>>8), byte(numVars&0xFF))
	for _, idx := range nameIdxs {
		c.bytecode = append(c.bytecode, byte(idx>>8), byte(idx&0xFF))
	}
}

// emitForIterEnterFast emits OpJSForIterEnterFast for non-capturing lexical loops.
func (c *Compiler) emitForIterEnterFast() {
	c.bytecode = append(c.bytecode, byte(OpJSForIterEnterFast))
}

// emitForIterExitFast emits OpJSForIterExitFast for non-capturing lexical loops.
func (c *Compiler) emitForIterExitFast() {
	c.bytecode = append(c.bytecode, byte(OpJSForIterExitFast))
}

// compileJScriptWithStatement compiles a non-strict ES5 with-statement.
func (c *Compiler) compileJScriptWithStatement(node *jsast.WithStatement) {
	if node == nil || node.Object == nil {
		return
	}

	// In strict mode, with statements are a syntax error
	if c.jsStrictMode {
		jsErr := jscript.NewJSSyntaxError(jscript.SyntaxError, 0, 0)
		jsErr.WithASPDescription("with statements are not allowed in strict mode")
		if c.sourceName != "" {
			jsErr.WithFile(c.sourceName)
		}
		panic(jsErr)
	}

	c.compileJScriptExpression(node.Object)
	c.emit(OpWithEnter)
	c.withDepth++
	if node.Body != nil {
		c.compileJScriptStatement(node.Body)
	}
	c.withDepth--
	c.emit(OpWithLeave)
}

// compileJScriptWhileStatement compiles: while (condition) statement
func (c *Compiler) compileJScriptWhileStatement(node *jsast.WhileStatement) {
	loopCtx := c.pushJSLoopContext()
	breakCtx := c.pushJSBreakContext()
	loopCtx.loopStart = len(c.bytecode)

	// Compile test condition
	c.compileJScriptExpression(node.Test)
	jumpExit := c.emitJSJump(OpJSJumpIfFalse)

	// Compile loop body
	if node.Body != nil {
		c.compileJScriptStatement(node.Body)
	}

	// Jump back to loop start
	c.emitJSJumpTo(OpJSJump, loopCtx.loopStart)

	// Patch exit jump
	c.patchJSJump(jumpExit)

	// Patch break and continue targets
	for _, breakPos := range breakCtx.breakTargets {
		c.patchJSJumpTo(breakPos, len(c.bytecode))
	}
	for _, contPos := range loopCtx.continueTargets {
		c.patchJSJumpTo(contPos, loopCtx.loopStart)
	}

	c.popJSLoopContext()
	c.popJSBreakContext()
}

// compileJScriptDoWhileStatement compiles: do statement while (condition)
func (c *Compiler) compileJScriptDoWhileStatement(node *jsast.DoWhileStatement) {
	loopCtx := c.pushJSLoopContext()
	breakCtx := c.pushJSBreakContext()
	loopCtx.loopStart = len(c.bytecode)

	// Compile loop body
	if node.Body != nil {
		c.compileJScriptStatement(node.Body)
	}

	// Mark continue target (test condition location)
	continueTarget := len(c.bytecode)

	// Compile test condition
	c.compileJScriptExpression(node.Test)

	// Jump back to loop if true
	c.emitJSJumpTo(OpJSJumpIfTrue, loopCtx.loopStart)

	// Patch break and continue targets
	for _, breakPos := range breakCtx.breakTargets {
		c.patchJSJumpTo(breakPos, len(c.bytecode))
	}
	for _, contPos := range loopCtx.continueTargets {
		c.patchJSJumpTo(contPos, continueTarget)
	}

	c.popJSLoopContext()
	c.popJSBreakContext()
}

// compileJScriptForStatement compiles: for (init; test; update) statement
func (c *Compiler) compileJScriptForStatement(node *jsast.ForStatement) {
	loopCtx := c.pushJSLoopContext()
	breakCtx := c.pushJSBreakContext()

	fastIntCounterName, fastIntLimitValue, fastIntEnabled := c.detectJSForFastIntLoop(node)
	fastIntCounterSlot := -1
	fastIntLimitSlot := -1
	if fastIntEnabled {
		c.jsPushLocalScope(false)
		counterHiddenName := fmt.Sprintf("__js_for_fast_counter__%d_%d", len(c.bytecode), c.jsLocalSlotCount)
		fastIntCounterSlot = c.jsDeclareCurrentLocal(counterHiddenName)
		if fastIntCounterSlot >= 0 && len(c.jsLocalScopeStack) > 0 {
			c.jsLocalScopeStack[len(c.jsLocalScopeStack)-1].entries[fastIntCounterName] = fastIntCounterSlot
			c.jsSetLocalType(fastIntCounterName, jsTypeInteger)
			limitHiddenName := fmt.Sprintf("__js_for_fast_limit__%d_%d", len(c.bytecode), c.jsLocalSlotCount)
			fastIntLimitSlot = c.jsDeclareCurrentLocal(limitHiddenName)
			if fastIntLimitSlot < 0 {
				fastIntEnabled = false
			}
		} else {
			fastIntEnabled = false
		}
		if !fastIntEnabled {
			c.jsPopLocalScope()
			fastIntCounterSlot = -1
			fastIntLimitSlot = -1
		}
	}

	// Fast path for `var` integer loops: `for (var i = N; i < M; i++)` or `<= M`.
	// The counter variable is stored in a local slot so all loop-variable accesses
	// (body and update) use OpJSGetLocal/OpJSSetLocal instead of the env hash map.
	// The variable persists after the loop with its final value (correct var scoping).
	fastIntVarCounterName := ""
	fastIntVarCounterSlot := -1
	fastIntVarLimitSlot := -1
	fastIntVarLimitValue := int64(0)
	fastIntVarEnabled := false
	if !fastIntEnabled {
		var varOK bool
		fastIntVarCounterName, fastIntVarLimitValue, varOK = c.detectJSForFastVarIntLoop(node)
		if varOK {
			// Obtain or create a local slot for the counter variable.
			// Prefer the existing slot from jsDeclareFunctionLocal (function scope).
			// Fall back to jsDeclareCurrentLocal for top-level blocks (root scope).
			counterSlot := c.jsDeclareFunctionLocal(fastIntVarCounterName)
			if counterSlot < 0 {
				counterSlot = c.jsDeclareCurrentLocal(fastIntVarCounterName)
			}
			if counterSlot >= 0 {
				fastIntVarCounterSlot = counterSlot
				c.jsSetLocalType(fastIntVarCounterName, jsTypeInteger)
				limitHiddenName := fmt.Sprintf("__js_for_fastvar_limit__%d_%d", len(c.bytecode), c.jsLocalSlotCount)
				fastIntVarLimitSlot = c.jsDeclareCurrentLocal(limitHiddenName)
				if fastIntVarLimitSlot < 0 {
					fastIntVarCounterSlot = -1
				} else {
					fastIntVarEnabled = true
				}
			}
		}
	}

	// Track whether we have a lexical (let/const) for-loop declaration
	var forIterNameIdxs []int
	var forIterNames []string
	isLexicalFor := false
	lexicalOuterScopeEntered := false
	forIterFastPath := false

	// Compile init expression
	if node.Initializer != nil {
		if fastIntEnabled {
			// let fast path: emit the integer initializer directly and store limit.
			if init, ok := node.Initializer.(*jsast.ForLoopInitializerLexicalDecl); ok && len(init.LexicalDeclaration.List) == 1 {
				binding := init.LexicalDeclaration.List[0]
				name, nameOK := jsBindingIdentifierName(binding.Target)
				_, initIntOK := jsNumericLiteralInt64(binding.Initializer)
				if nameOK && name == fastIntCounterName && binding.Initializer != nil && initIntOK {
					c.compileJScriptExpression(binding.Initializer)
					c.emit(OpJSSetLocal, fastIntCounterSlot)
					c.emit(OpConstant, c.addConstant(NewInteger(fastIntLimitValue)))
					c.emit(OpJSSetLocal, fastIntLimitSlot)
				} else {
					fastIntEnabled = false
					c.jsPopLocalScope()
					fastIntCounterSlot = -1
					fastIntLimitSlot = -1
				}
			} else {
				fastIntEnabled = false
				c.jsPopLocalScope()
				fastIntCounterSlot = -1
				fastIntLimitSlot = -1
			}
		}

		if !fastIntEnabled && fastIntVarEnabled {
			// var fast path: the counter slot was already declared by detection.
			// Emit the initializer into the counter slot and the limit into the hidden slot.
			if init, ok := node.Initializer.(*jsast.ForLoopInitializerVarDeclList); ok && len(init.List) == 1 {
				binding := init.List[0]
				name, nameOK := jsBindingIdentifierName(binding.Target)
				_, initIntOK := jsNumericLiteralInt64(binding.Initializer)
				if nameOK && name == fastIntVarCounterName && binding.Initializer != nil && initIntOK {
					c.compileJScriptExpression(binding.Initializer)
					c.emit(OpJSSetLocal, fastIntVarCounterSlot)
					c.emit(OpConstant, c.addConstant(NewInteger(fastIntVarLimitValue)))
					c.emit(OpJSSetLocal, fastIntVarLimitSlot)
				} else {
					fastIntVarEnabled = false
					fastIntVarCounterSlot = -1
					fastIntVarLimitSlot = -1
				}
			} else {
				fastIntVarEnabled = false
				fastIntVarCounterSlot = -1
				fastIntVarLimitSlot = -1
			}
		}

		if !fastIntEnabled && !fastIntVarEnabled {
			switch init := node.Initializer.(type) {
			case *jsast.ForLoopInitializerExpression:
				c.compileJScriptExpression(init.Expression)
				c.emit(OpJSPop)
			case *jsast.ForLoopInitializerVarDeclList:
				for _, binding := range init.List {
					name, ok := jsBindingIdentifierName(binding.Target)
					if !ok {
						continue
					}
					nameIdx := c.addConstant(NewString(name))
					localSlot := -1
					if c.jsLocalEnabled {
						localSlot = c.jsDeclareFunctionLocal(name)
					}
					if localSlot < 0 {
						c.emit(OpJSDeclareName, nameIdx)
					}
					if binding.Initializer != nil {
						c.compileJScriptExpression(binding.Initializer)
						if localSlot >= 0 {
							c.emit(OpJSSetLocal, localSlot)
						} else {
							c.emit(OpJSSetName, nameIdx)
						}
					}
				}
			case *jsast.ForLoopInitializerLexicalDecl:
				// ES6 let/const for-loop: create outer block scope for the loop variable
				lexDecl := init.LexicalDeclaration
				isLexicalFor = lexDecl.Token == jstoken.LET
				c.emit(OpJSBlockScopeEnter)
				lexicalOuterScopeEntered = true
				c.jsBlockScopeStack = append(c.jsBlockScopeStack, make(map[string]bool))
				for _, binding := range lexDecl.List {
					name, ok := jsBindingIdentifierName(binding.Target)
					if !ok {
						continue
					}
					nameIdx := c.addConstant(NewString(name))
					if lexDecl.Token == jstoken.CONST {
						c.emit(OpJSTDZRegisterConst, nameIdx)
						if binding.Initializer != nil {
							c.compileJScriptExpression(binding.Initializer)
							c.emitConstInitialize(nameIdx)
						}
					} else {
						c.emit(OpJSTDZRegisterLet, nameIdx)
						c.emit(OpJSLetDeclare, nameIdx)
						if binding.Initializer != nil {
							c.compileJScriptExpression(binding.Initializer)
							c.emit(OpJSSetName, nameIdx)
						}
					}
					if isLexicalFor {
						forIterNameIdxs = append(forIterNameIdxs, nameIdx)
						forIterNames = append(forIterNames, name)
					}
				}
				if isLexicalFor && len(forIterNames) > 0 {
					captureNames := make(map[string]struct{}, len(forIterNames))
					for _, n := range forIterNames {
						captureNames[n] = struct{}{}
					}
					forIterFastPath = !jsStatementCapturesLoopNames(node.Body, captureNames)
				}
			}
		}
	}

	loopCtx.loopStart = len(c.bytecode)

	// Compile test condition (jump out if false).
	// Use the fused OpJSJumpIfLessFast opcode for the extremely common pattern
	// `identifier < numericLiteral`, which covers virtually all ascending for-loops.
	// The fast opcode reads the variable directly from the JS environment and compares
	// it to the stored constant without touching the stack.
	var jumpExit int
	if node.Test != nil {
		if fastIntEnabled {
			c.bytecode = append(c.bytecode,
				byte(OpJSForFastIntEnter),
				byte(fastIntCounterSlot>>8), byte(fastIntCounterSlot),
				byte(fastIntLimitSlot>>8), byte(fastIntLimitSlot),
			)
			c.emit(OpJSGetLocal, fastIntCounterSlot)
			c.emit(OpJSGetLocal, fastIntLimitSlot)
			c.emit(OpJSLess)
			jumpExit = c.emitJSJump(OpJSJumpIfFalse)
		} else if fastIntVarEnabled {
			// var fast path: same fused structure as the let path.
			c.bytecode = append(c.bytecode,
				byte(OpJSForFastIntEnter),
				byte(fastIntVarCounterSlot>>8), byte(fastIntVarCounterSlot),
				byte(fastIntVarLimitSlot>>8), byte(fastIntVarLimitSlot),
			)
			c.emit(OpJSGetLocal, fastIntVarCounterSlot)
			c.emit(OpJSGetLocal, fastIntVarLimitSlot)
			c.emit(OpJSLess)
			jumpExit = c.emitJSJump(OpJSJumpIfFalse)
		} else {
			// Fused fast path for `identifier < numericLiteral` or `identifier <= numericLiteral`
			// when the identifier resolves to a local slot.  This avoids the full expression
			// compiler for the most common ascending-loop test patterns.
			if bin, ok := node.Test.(*jsast.BinaryExpression); ok &&
				(bin.Operator == jstoken.LESS || bin.Operator == jstoken.LESS_OR_EQUAL) {
				if id, ok := bin.Left.(*jsast.Identifier); ok {
					if slot, hasLocal := c.jsResolveLocalSlot(id.Name.String()); hasLocal {
						if num, isNum := bin.Right.(*jsast.NumberLiteral); isNum {
							c.emit(OpJSGetLocal, slot)
							switch v := num.Value.(type) {
							case int64:
								c.emit(OpConstant, c.addConstant(NewInteger(v)))
							case int:
								c.emit(OpConstant, c.addConstant(NewInteger(int64(v))))
							case float64:
								c.emit(OpConstant, c.addConstant(NewDouble(v)))
							default:
								c.compileJScriptExpression(node.Test)
								jumpExit = c.emitJSJump(OpJSJumpIfFalse)
								goto jsForTestDone
							}
							if bin.Operator == jstoken.LESS {
								c.emit(OpJSLess)
							} else {
								c.emit(OpJSLessEqual)
							}
							jumpExit = c.emitJSJump(OpJSJumpIfFalse)
							goto jsForTestDone
						}
					}
				}
			}
			if fastNameIdx, fastLimitIdx, ok := c.detectJSForFastLessTest(node.Test); ok {
				jumpExit = c.emitJSJumpIfLessFast(fastNameIdx, fastLimitIdx)
			} else {
				c.compileJScriptExpression(node.Test)
				jumpExit = c.emitJSJump(OpJSJumpIfFalse)
			}
		}
	jsForTestDone:
	}

	bodyStart := len(c.bytecode)

	// For let loops: enter per-iteration scope by copying loop vars into a child env frame
	if !fastIntEnabled && isLexicalFor && len(forIterNameIdxs) > 0 {
		if forIterFastPath {
			c.emitForIterEnterFast()
			c.jsForIterScopes = append(c.jsForIterScopes, jsForIterScope{fast: true})
		} else {
			c.emitForIterEnter(forIterNameIdxs)
			c.jsForIterScopes = append(c.jsForIterScopes, jsForIterScope{nameIdxs: forIterNameIdxs})
		}
	}

	// Compile loop body
	if node.Body != nil {
		c.compileJScriptStatement(node.Body)
	}
	if (fastIntEnabled || fastIntVarEnabled) && len(c.bytecode) == bodyStart {
		c.emit(OpNop)
	}

	// For let loops: exit per-iteration scope (write back updated vars to outer block scope).
	// This must be emitted BEFORE setting updateTarget so that continue statements (which
	// also emit ForIterExit via emitJSLeaveForIterScopes) jump to AFTER the ForIterExit
	// rather than triggering a double exit.
	if !fastIntEnabled && isLexicalFor && len(forIterNameIdxs) > 0 {
		c.jsForIterScopes = c.jsForIterScopes[:len(c.jsForIterScopes)-1]
		if forIterFastPath {
			c.emitForIterExitFast()
		} else {
			c.emitForIterExit(forIterNameIdxs)
		}
	}
	if fastIntEnabled {
		c.jsPopLocalScope()
	}
	// Note: no scope pop for fastIntVarEnabled — var remains visible after the loop.

	// Mark update target: continue statements jump here (after per-iteration scope exit).
	updateTarget := len(c.bytecode)

	// Compile update expression
	if fastIntEnabled {
		updateTarget = c.emitJSForFastInt(fastIntCounterSlot, fastIntLimitSlot, bodyStart)
	} else if fastIntVarEnabled {
		updateTarget = c.emitJSForFastInt(fastIntVarCounterSlot, fastIntVarLimitSlot, bodyStart)
	} else if node.Update != nil {
		handled, pushesResult := c.compileJScriptForUpdateFastPath(node.Update)
		if !handled {
			c.compileJScriptExpression(node.Update)
			c.emit(OpJSPop)
		} else if pushesResult {
			c.emit(OpJSPop)
		}
	}

	// Jump back to test
	if !fastIntEnabled && !fastIntVarEnabled {
		c.emitJSJumpTo(OpJSJump, loopCtx.loopStart)
	}

	// Patch exit jump
	if node.Test != nil {
		c.patchJSJump(jumpExit)
	}

	// Patch break targets to exit
	for _, breakPos := range breakCtx.breakTargets {
		c.patchJSJumpTo(breakPos, len(c.bytecode))
	}
	// Patch continue targets to updateTarget (per-iter exit + update)
	for _, contPos := range loopCtx.continueTargets {
		c.patchJSJumpTo(contPos, updateTarget)
	}

	c.popJSLoopContext()
	c.popJSBreakContext()

	// Exit outer block scope for lexical for-loop variables
	if lexicalOuterScopeEntered {
		if len(c.jsBlockScopeStack) > 0 {
			c.jsBlockScopeStack = c.jsBlockScopeStack[:len(c.jsBlockScopeStack)-1]
		}
		c.emit(OpJSBlockScopeExit)
	}
}

// compileJScriptForInStatement compiles: for (var in object) statement
func (c *Compiler) compileJScriptForInStatement(node *jsast.ForInStatement) {
	loopCtx := c.pushJSLoopContext()
	breakCtx := c.pushJSBreakContext()

	varName := ""
	declareName := false
	isLexical := false
	isConst := false
	switch into := node.Into.(type) {
	case *jsast.ForIntoVar:
		if into.Binding != nil {
			if name, ok := jsBindingIdentifierName(into.Binding.Target); ok {
				varName = name
				declareName = true
			}
		}
	case *jsast.ForDeclaration:
		// ES6 for (let/const k in obj)
		if target := into.Target; target != nil {
			if name, ok := jsBindingIdentifierName(target); ok {
				varName = name
				declareName = true
				isLexical = true
				isConst = into.IsConst
			}
		}
	case *jsast.ForIntoExpression:
		if id, ok := into.Expression.(*jsast.Identifier); ok {
			varName = id.Name.String()
		}
	}

	if varName == "" {
		c.popJSLoopContext()
		c.popJSBreakContext()
		return
	}

	c.compileJScriptExpression(node.Source)
	nameIdx := c.addConstant(NewString(varName))
	localForInSlot := -1
	if c.jsLocalEnabled && declareName && !isLexical {
		if slot, ok := c.jsResolveLocalSlot(varName); ok {
			localForInSlot = slot
		} else {
			localForInSlot = c.jsDeclareFunctionLocal(varName)
		}
	}

	// Lexical for-in: create outer block scope
	if isLexical {
		c.emit(OpJSBlockScopeEnter)
	}

	if declareName {
		if isConst {
			c.emit(OpJSTDZRegisterConst, nameIdx)
			// for-in initializes the variable on each iteration, so it will clear TDZ when entering iteration scope
		} else if isLexical {
			c.emit(OpJSTDZRegisterLet, nameIdx)
			c.emit(OpJSLetDeclare, nameIdx)
		} else {
			c.emit(OpJSDeclareName, nameIdx)
		}
	}

	loopCtx.loopStart = c.emitJSForIn(nameIdx)
	if localForInSlot >= 0 {
		// OpJSForIn assigns via jsSetName; mirror into local slot so loop bodies lowered
		// to OpJSGetLocal observe the current key on each iteration.
		c.emit(OpJSGetName, nameIdx)
		c.emit(OpJSSetLocal, localForInSlot)
	}

	if node.Body != nil {
		c.compileJScriptStatement(node.Body)
	}

	c.emitJSJumpTo(OpJSJump, loopCtx.loopStart)

	exitCleanup := len(c.bytecode)
	c.emit(OpJSForInCleanup, loopCtx.loopStart)
	exitPos := len(c.bytecode)

	c.patchJSForInExit(loopCtx.loopStart, exitPos)
	for _, breakPos := range breakCtx.breakTargets {
		c.patchJSJumpTo(breakPos, exitCleanup)
	}
	for _, contPos := range loopCtx.continueTargets {
		c.patchJSJumpTo(contPos, loopCtx.loopStart)
	}

	c.popJSLoopContext()
	c.popJSBreakContext()

	// Exit outer block scope for lexical for-in
	if isLexical {
		c.emit(OpJSBlockScopeExit)
	}
}

// emitJSForOf emits an OpJSForOf instruction with a placeholder exit target.
// Returns the bytecode position of the opcode so patchJSForOfExit can fill in the target.
func (c *Compiler) emitJSForOf(nameIdx int) int {
	pos := len(c.bytecode)
	c.bytecode = append(c.bytecode, byte(OpJSForOf), 0, 0, 0, 0, 0, 0)
	binary.BigEndian.PutUint16(c.bytecode[pos+1:], uint16(nameIdx))
	return pos
}

// patchJSForOfExit writes the resolved exit target into a previously emitted OpJSForOf.
func (c *Compiler) patchJSForOfExit(forOfPos int, target int) {
	if forOfPos < 0 || forOfPos+7 > len(c.bytecode) {
		panic("js for-of patch out of range")
	}
	binary.BigEndian.PutUint32(c.bytecode[forOfPos+3:], uint32(target))
}

// compileJScriptForOfStatement compiles: for (var/let/const x of iterable) statement
// It uses OpJSForOf which iterates values (not keys) similar to how OpJSForIn iterates keys.
func (c *Compiler) compileJScriptForOfStatement(node *jsast.ForOfStatement) {
	loopCtx := c.pushJSLoopContext()
	breakCtx := c.pushJSBreakContext()

	varName := ""
	declareName := false
	isLexical := false
	isConst := false
	switch into := node.Into.(type) {
	case *jsast.ForIntoVar:
		if into.Binding != nil {
			if name, ok := jsBindingIdentifierName(into.Binding.Target); ok {
				varName = name
				declareName = true
			}
		}
	case *jsast.ForDeclaration:
		// ES6 for (let/const x of iterable)
		if target := into.Target; target != nil {
			if name, ok := jsBindingIdentifierName(target); ok {
				varName = name
				declareName = true
				isLexical = true
				isConst = into.IsConst
			}
		}
	case *jsast.ForIntoExpression:
		if id, ok := into.Expression.(*jsast.Identifier); ok {
			varName = id.Name.String()
		}
	}

	if varName == "" {
		c.popJSLoopContext()
		c.popJSBreakContext()
		return
	}

	// Evaluate the iterable — its value is consumed by OpJSForOf on first entry.
	c.compileJScriptExpression(node.Source)
	nameIdx := c.addConstant(NewString(varName))

	// Lexical for-of: create an outer block scope so let/const is properly scoped.
	if isLexical {
		c.emit(OpJSBlockScopeEnter)
	}

	if declareName {
		if isConst {
			// For for-of, use let-style declaration so the loop header can re-initialize
			// the binding on each iteration without hitting TDZ or const-reassignment errors.
			// Per-iteration immutability is enforced at block scope; the value cannot be
			// reassigned inside the body because the let slot is still read-protected by the
			// enclosing block scope boundary.
			c.emit(OpJSTDZRegisterLet, nameIdx)
			c.emit(OpJSLetDeclare, nameIdx)
		} else if isLexical {
			c.emit(OpJSTDZRegisterLet, nameIdx)
			c.emit(OpJSLetDeclare, nameIdx)
		} else {
			c.emit(OpJSDeclareName, nameIdx)
		}
	}

	// Emit the loop header; loopStart is the position of OpJSForOf.
	loopCtx.loopStart = c.emitJSForOf(nameIdx)

	// Compile the loop body.
	if node.Body != nil {
		c.compileJScriptStatement(node.Body)
	}

	// Unconditional jump back to the OpJSForOf instruction to advance the iterator.
	c.emitJSJumpTo(OpJSJump, loopCtx.loopStart)

	// Emit cleanup opcode for early exits via break.
	exitCleanup := len(c.bytecode)
	c.bytecode = append(c.bytecode, byte(OpJSForOfCleanup), 0, 0, 0, 0)
	binary.BigEndian.PutUint32(c.bytecode[exitCleanup+1:], uint32(loopCtx.loopStart))
	exitPos := len(c.bytecode)

	// Patch the exit target in OpJSForOf to jump here when exhausted.
	c.patchJSForOfExit(loopCtx.loopStart, exitPos)

	// Patch break targets to the cleanup opcode.
	for _, breakPos := range breakCtx.breakTargets {
		c.patchJSJumpTo(breakPos, exitCleanup)
	}
	// Patch continue targets back to the OpJSForOf (advance to next value).
	for _, contPos := range loopCtx.continueTargets {
		c.patchJSJumpTo(contPos, loopCtx.loopStart)
	}

	c.popJSLoopContext()
	c.popJSBreakContext()

	// Exit the outer block scope for lexical for-of variables.
	if isLexical {
		c.emit(OpJSBlockScopeExit)
	}
}

// compileJScriptSwitchStatement compiles: switch (expr) { case ... default ... }
func (c *Compiler) compileJScriptSwitchStatement(node *jsast.SwitchStatement) {
	breakCtx := c.pushJSBreakContext()

	switchTmpName := fmt.Sprintf("__axonasp_js_switch_tmp_%d", len(c.bytecode))
	switchTmpIdx := c.addConstant(NewString(switchTmpName))

	c.emit(OpJSDeclareName, switchTmpIdx)
	c.compileJScriptExpression(node.Discriminant)
	c.emit(OpJSSetName, switchTmpIdx)

	caseBodyStart := make([]int, len(node.Body))
	caseMatchJumps := make([]int, 0, len(node.Body))

	for i := range node.Body {
		if node.Body[i] == nil || node.Body[i].Test == nil {
			continue
		}
		c.emit(OpJSGetName, switchTmpIdx)
		c.compileJScriptExpression(node.Body[i].Test)
		c.emit(OpJSStrictEq)
		jumpPos := c.emitJSJump(OpJSJumpIfTrue)
		caseMatchJumps = append(caseMatchJumps, jumpPos)
	}

	jumpToDefaultOrEnd := c.emitJSJump(OpJSJump)

	for i := range node.Body {
		caseBodyStart[i] = len(c.bytecode)
		if node.Body[i] == nil {
			continue
		}
		for j := range node.Body[i].Consequent {
			c.compileJScriptStatement(node.Body[i].Consequent[j])
		}
	}

	switchEnd := len(c.bytecode)

	matchIdx := 0
	for i := range node.Body {
		if node.Body[i] == nil || node.Body[i].Test == nil {
			continue
		}
		c.patchJSJumpTo(caseMatchJumps[matchIdx], caseBodyStart[i])
		matchIdx++
	}

	if node.Default >= 0 && node.Default < len(caseBodyStart) {
		c.patchJSJumpTo(jumpToDefaultOrEnd, caseBodyStart[node.Default])
	} else {
		c.patchJSJumpTo(jumpToDefaultOrEnd, switchEnd)
	}

	for _, breakPos := range breakCtx.breakTargets {
		c.patchJSJumpTo(breakPos, switchEnd)
	}

	c.popJSBreakContext()
}

// emitJSJumpTo emits an unconditional jump to a specific absolute target.
func (c *Compiler) emitJSJumpTo(op OpCode, target int) {
	c.emit(op, target)
}

// detectJSForFastLessTest checks whether a JScript for-loop test expression has the
// simple form `identifier < numericLiteral`.  If so, it interns the name and limit
// as constants and returns their indices along with true.  This pattern covers the
// overwhelming majority of ascending numeric for-loops.
func (c *Compiler) detectJSForFastLessTest(test jsast.Expression) (nameIdx, limitIdx int, ok bool) {
	bin, isBin := test.(*jsast.BinaryExpression)
	if !isBin || bin.Operator != jstoken.LESS {
		return 0, 0, false
	}
	id, isID := bin.Left.(*jsast.Identifier)
	if !isID {
		return 0, 0, false
	}
	num, isNum := bin.Right.(*jsast.NumberLiteral)
	if !isNum {
		return 0, 0, false
	}
	nameIdx = c.addConstant(NewString(id.Name.String()))
	switch v := num.Value.(type) {
	case int64:
		limitIdx = c.addConstant(NewInteger(v))
	case int:
		limitIdx = c.addConstant(NewInteger(int64(v)))
	case float64:
		limitIdx = c.addConstant(NewDouble(v))
	default:
		return 0, 0, false
	}
	return nameIdx, limitIdx, true
}

// emitJSJumpIfLessFast appends the 9-byte OpJSJumpIfLessFast instruction with a
// placeholder 4-byte exit target and returns the byte offset of that target for
// later patching via patchJSJump / patchJSJumpTo.
func (c *Compiler) emitJSJumpIfLessFast(nameIdx, limitIdx int) int {
	pos := len(c.bytecode)
	c.bytecode = append(c.bytecode,
		byte(OpJSJumpIfLessFast),
		byte(nameIdx>>8), byte(nameIdx),
		byte(limitIdx>>8), byte(limitIdx),
		0, 0, 0, 0, // placeholder exit target
	)
	// Return the byte offset of the 4-byte exit target for patchJSJump.
	return pos + 5
}

// emitJSMemberGet emits one IC-enabled JScript member-get opcode.
func (c *Compiler) emitJSMemberGet(nameIdx int) {
	c.emit(OpJSMemberGet, nameIdx)
}

// emitJSMemberSet emits one IC-enabled JScript member-set opcode.
func (c *Compiler) emitJSMemberSet(nameIdx int) {
	c.emit(OpJSMemberSet, nameIdx)
}

func (c *Compiler) emitJSForIn(nameIdx int) int {
	pos := len(c.bytecode)
	c.bytecode = append(c.bytecode, byte(OpJSForIn), 0, 0, 0, 0, 0, 0)
	binary.BigEndian.PutUint16(c.bytecode[pos+1:], uint16(nameIdx))
	return pos
}

func (c *Compiler) patchJSForInExit(forInPos int, target int) {
	if forInPos < 0 || forInPos+7 > len(c.bytecode) {
		panic("js for-in patch out of range")
	}
	binary.BigEndian.PutUint32(c.bytecode[forInPos+3:], uint32(target))
}

func (c *Compiler) compileJScriptExpression(expr jsast.Expression) {
	switch node := expr.(type) {
	case *jsast.NumberLiteral:
		switch v := node.Value.(type) {
		case int64:
			c.emit(OpConstant, c.addConstant(NewInteger(v)))
		case int:
			c.emit(OpConstant, c.addConstant(NewInteger(int64(v))))
		case float64:
			c.emit(OpConstant, c.addConstant(NewDouble(v)))
		case *big.Int:
			c.emit(OpConstant, c.addConstant(NewBigInt(v)))
		default:
			c.emit(OpConstant, c.addConstant(NewDouble(0)))
		}
	case *jsast.StringLiteral:
		c.emit(OpConstant, c.addConstant(NewString(node.Value.String())))
	case *jsast.BooleanLiteral:
		c.emit(OpConstant, c.addConstant(NewBool(node.Value)))
	case *jsast.NullLiteral:
		c.emit(OpConstant, c.addConstant(NewNull()))
	case *jsast.Identifier:
		// Check for VMENGINE global constant - returns AxonASP engine identification string.
		if node.Name.String() == "VMENGINE" {
			c.emit(JsOpAxonAsp)
		} else {
			if slot, ok := c.jsResolveLocalSlot(node.Name.String()); ok {
				c.emit(OpJSGetLocal, slot)
			} else {
				c.emit(OpJSGetName, c.addConstant(NewString(node.Name.String())))
			}
		}
	case *jsast.ThisExpression:
		c.emit(OpJSLoadThis)
	case *jsast.FunctionLiteral:
		c.compileJScriptFunctionLiteral(node, "", false)
	case *jsast.ClassExpression:
		c.compileJScriptClassLiteral(node.Class)
	case *jsast.BinaryExpression:
		switch node.Operator {
		case jstoken.LOGICAL_OR:
			c.compileJScriptExpression(node.Left)
			c.emit(OpJSDup)
			jumpTrue := c.emitJSJump(OpJSJumpIfTrue)
			c.emit(OpJSPop)
			c.compileJScriptExpression(node.Right)
			c.patchJSJump(jumpTrue)
		case jstoken.LOGICAL_AND:
			c.compileJScriptExpression(node.Left)
			c.emit(OpJSDup)
			jumpFalse := c.emitJSJump(OpJSJumpIfFalse)
			c.emit(OpJSPop)
			c.compileJScriptExpression(node.Right)
			c.patchJSJump(jumpFalse)
		case jstoken.COALESCE:
			c.compileJScriptExpression(node.Left)
			c.emit(OpJSDup)
			jumpNotNullish := c.emitJSJump(OpJSJumpIfNotNullish)
			c.emit(OpJSPop)
			c.compileJScriptExpression(node.Right)
			c.patchJSJump(jumpNotNullish)
		default:
			// Attempt compile-time constant folding on both operands, then on
			// the whole expression. If successful, emit a single OpConstant.
			foldedLeft := foldJSExpr(node.Left)
			foldedRight := foldJSExpr(node.Right)
			if folded := foldJSBinaryLiterals(node.Operator, foldedLeft, foldedRight); folded != nil {
				c.compileJScriptExpression(folded)
				return
			}

			// Phase 2: Optimize (x / 2) | 0 to x >> 1
			if node.Operator == jstoken.OR {
				if num, ok := foldedRight.(*jsast.NumberLiteral); ok {
					if v, ok := num.Value.(int64); ok && v == 0 {
						if bin, ok := foldedLeft.(*jsast.BinaryExpression); ok && bin.Operator == jstoken.SLASH {
							if num2, ok := bin.Right.(*jsast.NumberLiteral); ok {
								if v2, ok := num2.Value.(int64); ok && v2 == 2 {
									c.compileJScriptExpression(bin.Left)
									c.emit(OpConstant, c.addConstant(NewInteger(1)))
									c.emit(OpJSRightShift)
									return
								}
							}
						}
					}
				}
			}

			c.compileJScriptExpression(foldedLeft)
			c.compileJScriptExpression(foldedRight)
			switch node.Operator {
			case jstoken.PLUS:
				if c.jsInferredType(foldedLeft) == jsTypeInteger && c.jsInferredType(foldedRight) == jsTypeInteger {
					c.emit(OpJSAddInt)
				} else {
					c.emit(OpJSAdd)
				}
			case jstoken.MINUS:
				if c.jsInferredType(foldedLeft) == jsTypeInteger && c.jsInferredType(foldedRight) == jsTypeInteger {
					c.emit(OpJSSubInt)
				} else {
					c.emit(OpJSSubtract)
				}
			case jstoken.MULTIPLY:
				c.emit(OpJSMultiply)
			case jstoken.SLASH:
				c.emit(OpJSDivide)
			case jstoken.REMAINDER:
				c.emit(OpJSModulo)
			case jstoken.EXPONENT:
				c.emit(OpJSExponent)
			case jstoken.EQUAL:
				c.emit(OpJSLooseEqual)
			case jstoken.NOT_EQUAL:
				c.emit(OpJSLooseNotEqual)
			case jstoken.STRICT_EQUAL:
				c.emit(OpJSStrictEq)
			case jstoken.STRICT_NOT_EQUAL:
				c.emit(OpJSStrictNeq)
			case jstoken.LESS:
				c.emit(OpJSLess)
			case jstoken.GREATER:
				c.emit(OpJSGreater)
			case jstoken.LESS_OR_EQUAL:
				c.emit(OpJSLessEqual)
			case jstoken.GREATER_OR_EQUAL:
				c.emit(OpJSGreaterEqual)
			case jstoken.AND:
				c.emit(OpJSBitwiseAnd)
			case jstoken.OR:
				c.emit(OpJSBitwiseOr)
			case jstoken.EXCLUSIVE_OR:
				c.emit(OpJSBitwiseXor)
			case jstoken.SHIFT_LEFT:
				c.emit(OpJSLeftShift)
			case jstoken.SHIFT_RIGHT:
				c.emit(OpJSRightShift)
			case jstoken.UNSIGNED_SHIFT_RIGHT:
				c.emit(OpJSUnsignedRightShift)
			case jstoken.INSTANCEOF:
				c.emit(OpJSInstanceOf)
			case jstoken.IN:
				c.emit(OpJSIn)
			default:
				c.emit(OpJSLoadUndefined)
			}
		}
	case *jsast.AssignExpression:
		c.compileJScriptAssignment(node)
	case *jsast.PrivateDotExpression:
		c.compileJScriptExpression(node.Left)
		c.emitJSMemberGet(c.addConstant(NewString("\x00__priv_" + node.Identifier.Name.String())))
	case *jsast.DotExpression:
		if _, ok := node.Left.(*jsast.SuperExpression); ok {
			c.emit(OpJSSuperMemberGet, c.addConstant(NewString(node.Identifier.Name.String())))
			return
		}
		c.compileJScriptExpression(node.Left)
		c.emitJSMemberGet(c.addConstant(NewString(node.Identifier.Name.String())))
	case *jsast.BracketExpression:
		if _, ok := node.Left.(*jsast.SuperExpression); ok {
			c.compileJScriptExpression(node.Member)
			c.emit(OpJSSuperIndexGet)
			return
		}
		c.compileJScriptExpression(node.Left)
		c.compileJScriptExpression(node.Member)
		c.emit(OpJSIndexGet)
	case *jsast.ObjectLiteral:
		c.emit(OpJSNewObject)
		for i := 0; i < len(node.Value); i++ {
			switch prop := node.Value[i].(type) {
			case *jsast.PropertyShort:
				key := prop.Name.Name.String()
				c.emit(OpJSDup)
				if prop.Initializer != nil {
					c.compileJScriptExpression(prop.Initializer)
				} else {
					if slot, ok := c.jsResolveLocalSlot(key); ok {
						c.emit(OpJSGetLocal, slot)
					} else {
						c.emit(OpJSGetName, c.addConstant(NewString(key)))
					}
				}
				c.emitJSMemberSet(c.addConstant(NewString(key)))
			case *jsast.PropertyKeyed:
				if prop.Computed {
					// Computed property: { [expr]: value }
					// Stack before: ..., obj
					// Emit OpJSDup so we have a reference for OpJSComputedPropertySet.
					// Stack: ..., obj, obj (dup)
					c.emit(OpJSDup)
					// Compile value first so the dup is above it after key is pushed.
					c.compileJScriptExpression(prop.Value)
					// Stack: ..., obj, obj (dup), value
					c.compileJScriptExpression(prop.Key)
					// Stack: ..., obj, obj (dup), value, key
					// OpJSComputedPropertySet pops key, value, obj-dup → calls jsIndexSet(obj, key, value)
					// The outer obj reference remains on the stack.
					c.emit(OpJSComputedPropertySet)
					break
				}
				key, ok := jsObjectPropertyKeyName(prop.Key)
				if !ok {
					continue
				}
				c.emit(OpJSDup)
				c.compileJScriptExpression(prop.Value)
				switch prop.Kind {
				case jsast.PropertyKindGet:
					c.emitJSMemberSet(c.addConstant(NewString(jsAccessorGetterPrefix + key)))
				case jsast.PropertyKindSet:
					c.emitJSMemberSet(c.addConstant(NewString(jsAccessorSetterPrefix + key)))
				default:
					c.emitJSMemberSet(c.addConstant(NewString(key)))
				}
			}
		}
	case *jsast.ArrayLiteral:
		// ES6 spread in array literals is emitted as push/spread-push calls on a
		// single target array to preserve evaluation order.
		hasSpread := false
		for i := range node.Value {
			if _, ok := node.Value[i].(*jsast.SpreadElement); ok {
				hasSpread = true
				break
			}
		}
		if !hasSpread {
			for i := range node.Value {
				if node.Value[i] == nil {
					c.emit(OpJSLoadUndefined)
					continue
				}
				c.compileJScriptExpression(node.Value[i])
			}
			c.emit(OpJSNewArray, len(node.Value))
			break
		}

		c.emit(OpJSNewArray, 0)
		for i := range node.Value {
			c.emit(OpJSDup)
			if node.Value[i] == nil {
				c.emit(OpJSLoadUndefined)
				c.emit(OpJSCallMember, c.addConstant(NewString("push")), 1)
				c.emit(OpJSPop)
				continue
			}
			if spread, ok := node.Value[i].(*jsast.SpreadElement); ok {
				c.compileJScriptExpression(spread.Expression)
				c.emit(OpJSCallMember, c.addConstant(NewString("__spreadPush")), 1)
				c.emit(OpJSPop)
				continue
			}
			c.compileJScriptExpression(node.Value[i])
			c.emit(OpJSCallMember, c.addConstant(NewString("push")), 1)
			c.emit(OpJSPop)
		}
	case *jsast.RegExpLiteral:
		c.emit(OpJSGetName, c.addConstant(NewString("RegExp")))
		c.emit(OpConstant, c.addConstant(NewString(node.Pattern)))
		c.emit(OpConstant, c.addConstant(NewString(node.Flags)))
		c.emit(OpJSNew, 2)
	case *jsast.CallExpression:
		c.compileJScriptCall(node)
	case *jsast.NewExpression:
		c.compileJScriptExpression(node.Callee)
		for i := range node.ArgumentList {
			c.compileJScriptExpression(node.ArgumentList[i])
		}
		c.emit(OpJSNew, len(node.ArgumentList))
	case *jsast.UnaryExpression:
		if node.Operator == jstoken.TYPEOF {
			c.compileJScriptExpression(node.Operand)
			c.emit(OpJSTypeof)
			return
		}
		if node.Operator == jstoken.DELETE {
			switch t := node.Operand.(type) {
			case *jsast.DotExpression:
				c.compileJScriptExpression(t.Left)
				c.emit(OpJSMemberDelete, c.addConstant(NewString(t.Identifier.Name.String())))
			case *jsast.BracketExpression:
				c.compileJScriptExpression(t.Left)
				c.compileJScriptExpression(t.Member)
				c.emit(OpJSDelete)
			case *jsast.Identifier:
				// In strict mode, deleting a variable binding is a SyntaxError
				if c.jsStrictMode {
					jsErr := jscript.NewJSSyntaxError(jscript.IllegalAssignment, 0, 0)
					jsErr.WithASPDescription(fmt.Sprintf("cannot delete identifier '%s' in strict mode", t.Name.String()))
					if c.sourceName != "" {
						jsErr.WithFile(c.sourceName)
					}
					panic(jsErr)
				}
				// Non-strict mode: delete on an identifier always returns true (variables are not deletable)
				c.emit(OpConstant, c.addConstant(NewBool(true)))
			default:
				c.emit(OpConstant, c.addConstant(NewBool(true)))
			}
			return
		}
		if c.compileJScriptUpdateExpression(node) {
			return
		}
		switch node.Operator {
		case jstoken.MINUS:
			c.compileJScriptExpression(node.Operand)
			c.emit(OpJSNegate)
		case jstoken.NOT:
			c.compileJScriptExpression(node.Operand)
			c.emit(OpJSLogicalNot)
		case jstoken.BITWISE_NOT:
			c.compileJScriptExpression(node.Operand)
			c.emit(OpJSBitwiseNot)
		default:
			c.emit(OpJSLoadUndefined)
		}
	case *jsast.OptionalChain:
		prevCount := len(c.jsOptionalChainExits)
		c.compileJScriptExpression(node.Expression)
		if len(c.jsOptionalChainExits) > prevCount {
			jumpEnd := c.emitJSJump(OpJSJump)
			shortCircuitPos := len(c.bytecode)
			for i := prevCount; i < len(c.jsOptionalChainExits); i++ {
				c.patchJSJumpTo(c.jsOptionalChainExits[i], shortCircuitPos)
			}
			c.emit(OpJSPop)
			c.emit(OpJSLoadUndefined)
			c.patchJSJump(jumpEnd)
		}
		c.jsOptionalChainExits = c.jsOptionalChainExits[:prevCount]
	case *jsast.Optional:
		c.compileJScriptExpression(node.Expression)
		c.emit(OpJSDup)
		jumpPos := c.emitJSJump(OpJSJumpIfNullish)
		c.jsOptionalChainExits = append(c.jsOptionalChainExits, jumpPos)
	case *jsast.ConditionalExpression:
		c.compileJScriptExpression(node.Test)
		jumpFalse := c.emitJSJump(OpJSJumpIfFalse)
		c.compileJScriptExpression(node.Consequent)
		jumpEnd := c.emitJSJump(OpJSJump)
		c.patchJSJump(jumpFalse)
		c.compileJScriptExpression(node.Alternate)
		c.patchJSJump(jumpEnd)
	case *jsast.AwaitExpression:
		if node.Argument != nil {
			c.compileJScriptExpression(node.Argument)
		} else {
			c.emit(OpJSLoadUndefined)
		}
		c.emit(OpJSAwait)
	case *jsast.YieldExpression:
		if node.Argument != nil {
			c.compileJScriptExpression(node.Argument)
		} else {
			c.emit(OpJSLoadUndefined)
		}
		if node.Delegate {
			c.emit(OpJSYieldDelegate)
		} else {
			c.emit(OpJSYield)
		}
	case *jsast.TemplateLiteral:
		c.compileJScriptTemplateLiteral(node)
	case *jsast.ArrowFunctionLiteral:
		c.compileJScriptArrowFunctionLiteral(node)
	default:
		c.emit(OpJSLoadUndefined)
	}
}

func jsObjectPropertyKeyName(key jsast.Expression) (string, bool) {
	switch k := key.(type) {
	case *jsast.Identifier:
		return k.Name.String(), true
	case *jsast.StringLiteral:
		return k.Value.String(), true
	case *jsast.NumberLiteral:
		return k.Literal, true
	case *jsast.BooleanLiteral:
		if k.Value {
			return "true", true
		}
		return "false", true
	case *jsast.NullLiteral:
		return "null", true
	default:
		return "", false
	}
}

func (c *Compiler) compileJScriptAssignment(node *jsast.AssignExpression) {
	switch left := node.Left.(type) {
	case *jsast.Identifier:
		name := left.Name.String()
		// In strict mode, assigning to eval or arguments is a SyntaxError
		if c.jsStrictMode && jsIsRestrictedIdentifier(name) {
			jsErr := jscript.NewJSSyntaxError(jscript.IllegalAssignment, 0, 0)
			jsErr.WithASPDescription(fmt.Sprintf("cannot assign to '%s' in strict mode", name))
			if c.sourceName != "" {
				jsErr.WithFile(c.sourceName)
			}
			panic(jsErr)
		}
		nameIdx := c.addConstant(NewString(name))
		localSlot, hasLocal := c.jsResolveLocalSlot(name)
		if node.Operator == jstoken.ASSIGN {
			c.compileJScriptExpression(node.Right)
			if hasLocal && c.jsInferredType(node.Right) == jsTypeInteger {
				c.jsSetLocalType(name, jsTypeInteger)
			}
			c.emit(OpJSDup)
			if hasLocal {
				c.emit(OpJSSetLocal, localSlot)
			} else {
				c.emit(OpJSSetName, nameIdx)
			}
			return
		}
		if hasLocal {
			switch node.Operator {
			case jstoken.ADD_ASSIGN, jstoken.PLUS:
				c.emit(OpJSGetLocal, localSlot)
				c.compileJScriptExpression(node.Right)
				c.emit(OpJSAdd)
			case jstoken.SUBTRACT_ASSIGN, jstoken.MINUS:
				c.emit(OpJSGetLocal, localSlot)
				c.compileJScriptExpression(node.Right)
				c.emit(OpJSSubtract)
			case jstoken.MULTIPLY_ASSIGN, jstoken.MULTIPLY:
				c.emit(OpJSGetLocal, localSlot)
				c.compileJScriptExpression(node.Right)
				c.emit(OpJSMultiply)
			case jstoken.QUOTIENT_ASSIGN, jstoken.SLASH:
				c.emit(OpJSGetLocal, localSlot)
				c.compileJScriptExpression(node.Right)
				c.emit(OpJSDivide)
			case jstoken.REMAINDER_ASSIGN, jstoken.REMAINDER:
				c.emit(OpJSGetLocal, localSlot)
				c.compileJScriptExpression(node.Right)
				c.emit(OpJSModulo)
			case jstoken.EXPONENT_ASSIGN, jstoken.EXPONENT:
				c.emit(OpJSGetLocal, localSlot)
				c.compileJScriptExpression(node.Right)
				c.emit(OpJSExponent)
			default:
				c.compileJScriptExpression(node.Right)
				c.emit(OpJSSetName, nameIdx)
				return
			}
			c.emit(OpJSDup)
			c.emit(OpJSSetLocal, localSlot)
			c.emit(OpJSLoadUndefined)
			return
		}
		c.compileJScriptExpression(node.Right)
		switch node.Operator {
		case jstoken.ADD_ASSIGN, jstoken.PLUS:
			c.emit(OpJSAddAssign, nameIdx)
		case jstoken.SUBTRACT_ASSIGN, jstoken.MINUS:
			c.emit(OpJSSubtractAssign, nameIdx)
		case jstoken.MULTIPLY_ASSIGN, jstoken.MULTIPLY:
			c.emit(OpJSMultiplyAssign, nameIdx)
		case jstoken.QUOTIENT_ASSIGN, jstoken.SLASH:
			c.emit(OpJSDivideAssign, nameIdx)
		case jstoken.REMAINDER_ASSIGN, jstoken.REMAINDER:
			c.emit(OpJSModuloAssign, nameIdx)
		case jstoken.EXPONENT_ASSIGN, jstoken.EXPONENT:
			c.emit(OpJSExponentAssign, nameIdx)
			return
		case jstoken.LOGICAL_AND_ASSIGN, jstoken.LOGICAL_AND:
			c.emit(OpJSLogicalAndAssign, nameIdx)
			return
		case jstoken.LOGICAL_OR_ASSIGN, jstoken.LOGICAL_OR:
			c.emit(OpJSLogicalOrAssign, nameIdx)
			return
		case jstoken.COALESCE_ASSIGN, jstoken.COALESCE:
			c.emit(OpJSCoalesceAssign, nameIdx)
			return
		default:
			c.emit(OpJSSetName, nameIdx)
		}
		// Compound assignments in AxonASP currently don't return the value on stack after OpJSXXXAssign?
		// Let's check OpJSAddAssign etc.
		c.emit(OpJSLoadUndefined)
	case *jsast.ObjectPattern, *jsast.ArrayPattern:
		if node.Operator != jstoken.ASSIGN {
			jsErr := jscript.NewJSSyntaxError(jscript.IllegalAssignment, 0, 0)
			if c.sourceName != "" {
				jsErr.WithFile(c.sourceName)
			}
			panic(jsErr)
		}
		c.compileJScriptExpression(node.Right)
		c.emit(OpJSDup)
		c.compileJScriptDestructuring(left.(jsast.BindingTarget), false, false, false)
	case *jsast.PrivateDotExpression:
		c.compileJScriptExpression(left.Left)
		c.compileJScriptExpression(node.Right)
		c.emitJSMemberSet(c.addConstant(NewString("\x00__priv_" + left.Identifier.Name.String())))
		c.emit(OpJSLoadUndefined)
	case *jsast.DotExpression:
		if _, ok := left.Left.(*jsast.SuperExpression); ok {
			c.compileJScriptExpression(node.Right)
			c.emit(OpJSSuperMemberSet, c.addConstant(NewString(left.Identifier.Name.String())))
			return
		}
		c.compileJScriptExpression(left.Left)
		c.compileJScriptExpression(node.Right)
		c.emitJSMemberSet(c.addConstant(NewString(left.Identifier.Name.String())))
		c.emit(OpJSLoadUndefined)
	case *jsast.BracketExpression:
		if _, ok := left.Left.(*jsast.SuperExpression); ok {
			c.compileJScriptExpression(node.Right)
			c.compileJScriptExpression(left.Member)
			c.emit(OpJSSuperIndexSet)
			return
		}
		c.compileJScriptExpression(node.Right)
		c.compileJScriptExpression(left.Left)
		c.compileJScriptExpression(left.Member)
		c.emit(OpJSIndexSet)
		c.emit(OpJSLoadUndefined)
	case *jsast.CallExpression:
		switch callee := left.Callee.(type) {
		case *jsast.Identifier:
			c.emit(OpJSGetName, c.addConstant(NewString(callee.Name.String())))
			for i := range left.ArgumentList {
				c.compileJScriptExpression(left.ArgumentList[i])
			}
			c.compileJScriptExpression(node.Right)
			c.emit(OpJSCall, len(left.ArgumentList)+1)
			c.emit(OpJSPop)
			c.emit(OpJSLoadUndefined)
		case *jsast.DotExpression:
			c.compileJScriptExpression(callee.Left)
			for i := range left.ArgumentList {
				c.compileJScriptExpression(left.ArgumentList[i])
			}
			c.compileJScriptExpression(node.Right)
			c.emit(OpJSCallMember, c.addConstant(NewString(callee.Identifier.Name.String())), len(left.ArgumentList)+1)
			c.emit(OpJSPop)
			c.emit(OpJSLoadUndefined)
		default:
			c.emit(OpJSLoadUndefined)
		}
	default:
		c.emit(OpJSLoadUndefined)
	}
}

// compileJScriptForUpdateFastPath emits optimized update bytecode for common loop
// forms that increment or decrement one identifier by one.
// Return values are (handled, pushesResult).
func (c *Compiler) compileJScriptForUpdateFastPath(expr jsast.Expression) (bool, bool) {
	if update, ok := expr.(*jsast.UnaryExpression); ok {
		if ident, ok := update.Operand.(*jsast.Identifier); ok {
			if slot, hasLocal := c.jsResolveLocalSlot(ident.Name.String()); hasLocal {
				switch update.Operator {
				case jstoken.INCREMENT:
					c.emit(OpJSIncLocal, slot)
					return true, false
				case jstoken.DECREMENT:
					c.emit(OpJSDecLocal, slot)
					return true, false
				}
			}
			nameIdx := c.addConstant(NewString(ident.Name.String()))
			switch update.Operator {
			case jstoken.INCREMENT:
				c.emit(OpJSIncLocalInt, nameIdx)
				return true, false
			case jstoken.DECREMENT:
				c.emit(OpJSDecLocalInt, nameIdx)
				return true, false
			}
		}
	}

	assign, ok := expr.(*jsast.AssignExpression)
	if !ok {
		return false, false
	}

	leftID, ok := assign.Left.(*jsast.Identifier)
	if !ok {
		return false, false
	}

	name := leftID.Name.String()
	nameIdx := c.addConstant(NewString(name))
	if slot, hasLocal := c.jsResolveLocalSlot(name); hasLocal {
		if assign.Operator == jstoken.ADD_ASSIGN || assign.Operator == jstoken.SUBTRACT_ASSIGN {
			if !jsIsNumericOneLiteral(assign.Right) {
				return false, false
			}
			if assign.Operator == jstoken.ADD_ASSIGN {
				c.emit(OpJSIncLocal, slot)
			} else {
				c.emit(OpJSDecLocal, slot)
			}
			return true, false
		}

		if assign.Operator == jstoken.ASSIGN {
			rightBin, ok := assign.Right.(*jsast.BinaryExpression)
			if !ok {
				return false, false
			}
			rightLeftID, ok := rightBin.Left.(*jsast.Identifier)
			if !ok || rightLeftID.Name.String() != name {
				return false, false
			}
			if !jsIsNumericOneLiteral(rightBin.Right) {
				return false, false
			}
			if rightBin.Operator == jstoken.PLUS {
				c.emit(OpJSIncLocal, slot)
				return true, false
			}
			if rightBin.Operator == jstoken.MINUS {
				c.emit(OpJSDecLocal, slot)
				return true, false
			}
		}
	}

	// i += 1 / i -= 1
	if assign.Operator == jstoken.ADD_ASSIGN || assign.Operator == jstoken.SUBTRACT_ASSIGN {
		if !jsIsNumericOneLiteral(assign.Right) {
			return false, false
		}
		if assign.Operator == jstoken.ADD_ASSIGN {
			c.emit(OpJSIncLocalInt, nameIdx)
		} else {
			c.emit(OpJSDecLocalInt, nameIdx)
		}
		return true, false
	}

	// i = i + 1 / i = i - 1
	if assign.Operator != jstoken.ASSIGN {
		return false, false
	}

	rightBin, ok := assign.Right.(*jsast.BinaryExpression)
	if !ok {
		return false, false
	}
	rightLeftID, ok := rightBin.Left.(*jsast.Identifier)
	if !ok || rightLeftID.Name.String() != name {
		return false, false
	}
	if !jsIsNumericOneLiteral(rightBin.Right) {
		return false, false
	}

	if rightBin.Operator == jstoken.PLUS {
		c.emit(OpJSIncLocalInt, nameIdx)
		return true, false
	}
	if rightBin.Operator == jstoken.MINUS {
		c.emit(OpJSDecLocalInt, nameIdx)
		return true, false
	}

	return false, false
}

// jsIsNumericOneLiteral returns true when expr is a numeric literal with value 1.
func jsIsNumericOneLiteral(expr jsast.Expression) bool {
	num, ok := expr.(*jsast.NumberLiteral)
	if !ok {
		return false
	}
	switch v := num.Value.(type) {
	case int:
		return v == 1
	case int32:
		return v == 1
	case int64:
		return v == 1
	case float32:
		return v == 1
	case float64:
		return v == 1
	default:
		return false
	}
}

// jsNumericLiteralInt64 returns the integer value of a numeric literal when it is
// representable as an int64 without loss.
func jsNumericLiteralInt64(expr jsast.Expression) (int64, bool) {
	num, ok := expr.(*jsast.NumberLiteral)
	if !ok {
		return 0, false
	}
	switch v := num.Value.(type) {
	case int:
		return int64(v), true
	case int32:
		return int64(v), true
	case int64:
		return v, true
	case float32:
		if float64(int64(v)) == float64(v) {
			return int64(v), true
		}
	case float64:
		if math.Trunc(v) == v {
			return int64(v), true
		}
	case *big.Int:
		if v.IsInt64() {
			return v.Int64(), true
		}
	}
	return 0, false
}

// jsIsNumericZeroLiteral returns true when expr is a numeric literal with value 0.
func jsIsNumericZeroLiteral(expr jsast.Expression) bool {
	value, ok := jsNumericLiteralInt64(expr)
	return ok && value == 0
}

// detectJSForFastIntLoop checks whether a JScript for-loop matches the fast integer
// loop shape: `for (let i = N; i < M; i++)` or `for (let i = N; i <= M; i++)`.
// N can be any integer literal. For `<=`, limitValue is returned as M+1 (exclusive
// upper bound) so the caller stores it and uses `<` semantics via OpJSForFastInt.
// Returns (counterName, limitValue, ok) where limitValue is the exclusive upper bound.
func (c *Compiler) detectJSForFastIntLoop(node *jsast.ForStatement) (counterName string, limitValue int64, ok bool) {
	if node == nil || !c.jsLocalEnabled || node.Initializer == nil || node.Test == nil || node.Update == nil {
		return "", 0, false
	}
	if jsStatementContainsNestedFunction(node.Body) {
		return "", 0, false
	}
	init, isLexical := node.Initializer.(*jsast.ForLoopInitializerLexicalDecl)
	if !isLexical || init.LexicalDeclaration.Token != jstoken.LET {
		return "", 0, false
	}
	if len(init.LexicalDeclaration.List) != 1 {
		return "", 0, false
	}
	binding := init.LexicalDeclaration.List[0]
	name, isName := jsBindingIdentifierName(binding.Target)
	// Accept any integer literal as the initial value (not only 0).
	if !isName || binding.Initializer == nil {
		return "", 0, false
	}
	if _, initOK := jsNumericLiteralInt64(binding.Initializer); !initOK {
		return "", 0, false
	}
	bin, isBin := node.Test.(*jsast.BinaryExpression)
	// Accept both `<` and `<=` test conditions.
	if !isBin || (bin.Operator != jstoken.LESS && bin.Operator != jstoken.LESS_OR_EQUAL) {
		return "", 0, false
	}
	leftID, isLeftID := bin.Left.(*jsast.Identifier)
	if !isLeftID || leftID.Name.String() != name {
		return "", 0, false
	}
	limit, isLimit := jsNumericLiteralInt64(bin.Right)
	if !isLimit {
		return "", 0, false
	}
	// For `<=`, convert to exclusive upper bound (i <= N → i < N+1).
	if bin.Operator == jstoken.LESS_OR_EQUAL {
		if limit == math.MaxInt64 {
			// Overflow guard: cannot represent N+1, fall back to generic path.
			return "", 0, false
		}
		limit++
	}
	update, isUpdate := node.Update.(*jsast.UnaryExpression)
	if !isUpdate || update.Operator != jstoken.INCREMENT {
		return "", 0, false
	}
	updateID, isUpdateID := update.Operand.(*jsast.Identifier)
	if !isUpdateID || updateID.Name.String() != name {
		return "", 0, false
	}
	return name, limit, true
}

// detectJSForFastVarIntLoop checks whether a JScript for-loop using a `var`
// declaration matches a fast integer loop shape:
// `for (var i = N; i < M; i++)` or `for (var i = N; i <= M; i++)`
// where N and M are integer literals and jsLocalEnabled is active.
// Unlike detectJSForFastIntLoop (which handles `let`), this path reuses the
// function-local or root-scope slot for the counter so the value persists
// after the loop (correct `var` scoping behaviour).
// Returns (counterName, limitValue, ok) where limitValue is the exclusive upper bound.
// Slot allocation happens in the compiler, not here, to avoid early side-effects.
func (c *Compiler) detectJSForFastVarIntLoop(node *jsast.ForStatement) (counterName string, limitValue int64, ok bool) {
	if node == nil || !c.jsLocalEnabled || node.Initializer == nil || node.Test == nil || node.Update == nil {
		return "", 0, false
	}
	if jsStatementContainsNestedFunction(node.Body) {
		return "", 0, false
	}
	init, isVarDecl := node.Initializer.(*jsast.ForLoopInitializerVarDeclList)
	if !isVarDecl || len(init.List) != 1 {
		return "", 0, false
	}
	binding := init.List[0]
	name, isName := jsBindingIdentifierName(binding.Target)
	if !isName || binding.Initializer == nil {
		return "", 0, false
	}
	if _, initOK := jsNumericLiteralInt64(binding.Initializer); !initOK {
		return "", 0, false
	}
	bin, isBin := node.Test.(*jsast.BinaryExpression)
	if !isBin || (bin.Operator != jstoken.LESS && bin.Operator != jstoken.LESS_OR_EQUAL) {
		return "", 0, false
	}
	leftID, isLeftID := bin.Left.(*jsast.Identifier)
	if !isLeftID || leftID.Name.String() != name {
		return "", 0, false
	}
	limit, isLimit := jsNumericLiteralInt64(bin.Right)
	if !isLimit {
		return "", 0, false
	}
	// For `<=`, convert to exclusive upper bound (i <= N → i < N+1).
	if bin.Operator == jstoken.LESS_OR_EQUAL {
		if limit == math.MaxInt64 {
			return "", 0, false
		}
		limit++
	}
	update, isUpdate := node.Update.(*jsast.UnaryExpression)
	if !isUpdate || update.Operator != jstoken.INCREMENT {
		return "", 0, false
	}
	updateID, isUpdateID := update.Operand.(*jsast.Identifier)
	if !isUpdateID || updateID.Name.String() != name {
		return "", 0, false
	}
	return name, limit, true
}

// emitJSForFastInt appends the fused JScript integer for-loop opcode and returns the
// byte offset of the opcode start so continue targets can jump to it.
// The encoded jump operand is a relative back-jump distance from the instruction end.
func (c *Compiler) emitJSForFastInt(counterSlot, limitSlot, bodyTarget int) int {
	pos := len(c.bytecode)
	jumpOffset := (pos + 9) - bodyTarget
	if jumpOffset < 0 {
		jumpOffset = 0
	}
	c.bytecode = append(c.bytecode,
		byte(OpJSForFastInt),
		byte(counterSlot>>8), byte(counterSlot),
		byte(limitSlot>>8), byte(limitSlot),
		byte(jumpOffset>>24), byte(jumpOffset>>16), byte(jumpOffset>>8), byte(jumpOffset),
	)
	return pos
}

func (c *Compiler) compileJScriptCall(node *jsast.CallExpression) {
	switch callee := node.Callee.(type) {
	case *jsast.SuperExpression:
		if !c.jsIsDerivedConstructor {
			jsErr := jscript.NewJSSyntaxError(jscript.SyntaxError, 0, 0)
			jsErr.WithASPDescription("super() keyword unexpected here")
			if c.sourceName != "" {
				jsErr.WithFile(c.sourceName)
			}
			panic(jsErr)
		}
		for i := range node.ArgumentList {
			c.compileJScriptExpression(node.ArgumentList[i])
		}
		c.emit(OpJSSuperCall, len(node.ArgumentList))
		c.emit(OpJSDup)
		c.emit(OpJSSetThis)

		// Inject field initialization for derived class constructors
		c.compileJScriptClassFields()
	case *jsast.DotExpression:
		if _, ok := callee.Left.(*jsast.SuperExpression); ok {
			// super.method(...)
			for i := range node.ArgumentList {
				c.compileJScriptExpression(node.ArgumentList[i])
			}
			c.emit(OpJSSuperCallMember, c.addConstant(NewString(callee.Identifier.Name.String())), len(node.ArgumentList))
			return
		}

		// Phase 1: Math Object Interception & Phase 2: Math.floor(x / 2) optimization
		if id, ok := callee.Left.(*jsast.Identifier); ok && id.Name.String() == "Math" {
			method := callee.Identifier.Name.String()
			switch method {
			case "sin", "cos", "tan", "abs", "floor", "ceil", "round", "sqrt":
				if len(node.ArgumentList) == 1 {
					// Phase 2: Optimize Math.floor(x / 2) to x >> 1
					if method == "floor" {
						if bin, ok := node.ArgumentList[0].(*jsast.BinaryExpression); ok && bin.Operator == jstoken.SLASH {
							if num, ok := bin.Right.(*jsast.NumberLiteral); ok {
								if v, ok := num.Value.(int64); ok && v == 2 {
									c.compileJScriptExpression(bin.Left)
									c.emit(OpConstant, c.addConstant(NewInteger(1)))
									c.emit(OpJSRightShift)
									return
								}
							}
						}
					}

					c.compileJScriptExpression(node.ArgumentList[0])
					switch method {
					case "sin":
						c.emit(OpJSMathSin)
					case "cos":
						c.emit(OpJSMathCos)
					case "tan":
						c.emit(OpJSMathTan)
					case "abs":
						c.emit(OpJSMathAbs)
					case "floor":
						c.emit(OpJSMathFloor)
					case "ceil":
						c.emit(OpJSMathCeil)
					case "round":
						c.emit(OpJSMathRound)
					case "sqrt":
						c.emit(OpJSMathSqrt)
					}
					return
				}
			case "min", "max":
				if len(node.ArgumentList) == 2 {
					c.compileJScriptExpression(node.ArgumentList[0])
					c.compileJScriptExpression(node.ArgumentList[1])
					if method == "min" {
						c.emit(OpJSMathMin)
					} else {
						c.emit(OpJSMathMax)
					}
					return
				}
			}
		}

		c.compileJScriptExpression(callee.Left)
		for i := range node.ArgumentList {
			c.compileJScriptExpression(node.ArgumentList[i])
		}
		c.emit(OpJSCallMember, c.addConstant(NewString(callee.Identifier.Name.String())), len(node.ArgumentList))
	case *jsast.BracketExpression:
		if _, ok := callee.Left.(*jsast.SuperExpression); ok {
			// super[index](...)
			for i := range node.ArgumentList {
				c.compileJScriptExpression(node.ArgumentList[i])
			}
			c.compileJScriptExpression(callee.Member)
			c.emit(OpJSSuperCallComputedMember, len(node.ArgumentList))
			return
		}
		c.compileJScriptExpression(callee.Left)
		c.compileJScriptExpression(callee.Member)
		for i := range node.ArgumentList {
			c.compileJScriptExpression(node.ArgumentList[i])
		}
		c.emit(OpJSCallComputedMember, len(node.ArgumentList))
	default:
		c.compileJScriptExpression(node.Callee)
		for i := range node.ArgumentList {
			c.compileJScriptExpression(node.ArgumentList[i])
		}
		c.emit(OpJSCall, len(node.ArgumentList))
	}
}

// compileJScriptTailReturn emits one tail-call opcode when one return argument is a call expression.
func (c *Compiler) compileJScriptTailReturn(argument jsast.Expression) bool {
	if !c.jsStrictMode {
		return false
	}
	callExpr, ok := argument.(*jsast.CallExpression)
	if !ok {
		return false
	}

	switch callee := callExpr.Callee.(type) {
	case *jsast.DotExpression:
		c.compileJScriptExpression(callee.Left)
		for i := range callExpr.ArgumentList {
			c.compileJScriptExpression(callExpr.ArgumentList[i])
		}
		c.emit(OpJSTailCallMember, c.addConstant(NewString(callee.Identifier.Name.String())), len(callExpr.ArgumentList))
	case *jsast.BracketExpression:
		c.compileJScriptExpression(callee.Left)
		c.compileJScriptExpression(callee.Member)
		for i := range callExpr.ArgumentList {
			c.compileJScriptExpression(callExpr.ArgumentList[i])
		}
		c.emit(OpJSTailCallComputedMember, len(callExpr.ArgumentList))
	default:
		c.compileJScriptExpression(callExpr.Callee)
		for i := range callExpr.ArgumentList {
			c.compileJScriptExpression(callExpr.ArgumentList[i])
		}
		c.emit(OpJSTailCall, len(callExpr.ArgumentList))
	}

	return true
}

func jsFunctionContainsNestedFunction(fn *jsast.FunctionLiteral) bool {
	if fn == nil || fn.Body == nil {
		return false
	}
	for i := range fn.Body.List {
		if jsStatementContainsNestedFunction(fn.Body.List[i]) {
			return true
		}
	}
	return false
}

func jsProgramContainsNestedFunction(stmts []jsast.Statement) bool {
	for i := range stmts {
		if jsStatementContainsNestedFunction(stmts[i]) {
			return true
		}
	}
	return false
}

func jsStatementContainsNestedFunction(stmt jsast.Statement) bool {
	if stmt == nil {
		return false
	}
	switch node := stmt.(type) {
	case *jsast.FunctionDeclaration:
		return true
	case *jsast.BlockStatement:
		for i := range node.List {
			if jsStatementContainsNestedFunction(node.List[i]) {
				return true
			}
		}
	case *jsast.ExpressionStatement:
		return jsExpressionContainsNestedFunction(node.Expression)
	case *jsast.IfStatement:
		if jsExpressionContainsNestedFunction(node.Test) {
			return true
		}
		if jsStatementContainsNestedFunction(node.Consequent) {
			return true
		}
		if jsStatementContainsNestedFunction(node.Alternate) {
			return true
		}
	case *jsast.ForStatement:
		if node.Initializer != nil {
			switch init := node.Initializer.(type) {
			case *jsast.ForLoopInitializerExpression:
				if jsExpressionContainsNestedFunction(init.Expression) {
					return true
				}
			case *jsast.ForLoopInitializerVarDeclList:
				for _, b := range init.List {
					if b != nil && jsExpressionContainsNestedFunction(b.Initializer) {
						return true
					}
				}
			case *jsast.ForLoopInitializerLexicalDecl:
				for _, b := range init.LexicalDeclaration.List {
					if b != nil && jsExpressionContainsNestedFunction(b.Initializer) {
						return true
					}
				}
			}
		}
		if jsExpressionContainsNestedFunction(node.Test) || jsExpressionContainsNestedFunction(node.Update) {
			return true
		}
		return jsStatementContainsNestedFunction(node.Body)
	case *jsast.ReturnStatement:
		return jsExpressionContainsNestedFunction(node.Argument)
	case *jsast.ThrowStatement:
		return jsExpressionContainsNestedFunction(node.Argument)
	case *jsast.WhileStatement:
		return jsExpressionContainsNestedFunction(node.Test) || jsStatementContainsNestedFunction(node.Body)
	case *jsast.DoWhileStatement:
		return jsExpressionContainsNestedFunction(node.Test) || jsStatementContainsNestedFunction(node.Body)
	case *jsast.SwitchStatement:
		if jsExpressionContainsNestedFunction(node.Discriminant) {
			return true
		}
		for i := range node.Body {
			clause := node.Body[i]
			if clause == nil {
				continue
			}
			if jsExpressionContainsNestedFunction(clause.Test) {
				return true
			}
			for j := range clause.Consequent {
				if jsStatementContainsNestedFunction(clause.Consequent[j]) {
					return true
				}
			}
		}
	case *jsast.TryStatement:
		if jsStatementContainsNestedFunction(node.Body) {
			return true
		}
		if node.Catch != nil && jsStatementContainsNestedFunction(node.Catch.Body) {
			return true
		}
		if node.Finally != nil && jsStatementContainsNestedFunction(node.Finally) {
			return true
		}
	case *jsast.UsingDeclaration:
		for _, b := range node.List {
			if b != nil && jsExpressionContainsNestedFunction(b.Initializer) {
				return true
			}
		}
	}
	return false
}

func jsExpressionContainsNestedFunction(expr jsast.Expression) bool {
	if expr == nil {
		return false
	}
	switch node := expr.(type) {
	case *jsast.FunctionLiteral, *jsast.ArrowFunctionLiteral:
		return true
	case *jsast.AssignExpression:
		return jsExpressionContainsNestedFunction(node.Left) || jsExpressionContainsNestedFunction(node.Right)
	case *jsast.BinaryExpression:
		return jsExpressionContainsNestedFunction(node.Left) || jsExpressionContainsNestedFunction(node.Right)
	case *jsast.UnaryExpression:
		return jsExpressionContainsNestedFunction(node.Operand)
	case *jsast.DotExpression:
		return jsExpressionContainsNestedFunction(node.Left)
	case *jsast.BracketExpression:
		return jsExpressionContainsNestedFunction(node.Left) || jsExpressionContainsNestedFunction(node.Member)
	case *jsast.CallExpression:
		if jsExpressionContainsNestedFunction(node.Callee) {
			return true
		}
		for i := range node.ArgumentList {
			if jsExpressionContainsNestedFunction(node.ArgumentList[i]) {
				return true
			}
		}
	case *jsast.NewExpression:
		if jsExpressionContainsNestedFunction(node.Callee) {
			return true
		}
		for i := range node.ArgumentList {
			if jsExpressionContainsNestedFunction(node.ArgumentList[i]) {
				return true
			}
		}
	case *jsast.ObjectLiteral:
		for i := range node.Value {
			switch p := node.Value[i].(type) {
			case *jsast.PropertyShort:
				if jsExpressionContainsNestedFunction(p.Initializer) {
					return true
				}
			case *jsast.PropertyKeyed:
				if jsExpressionContainsNestedFunction(p.Key) || jsExpressionContainsNestedFunction(p.Value) {
					return true
				}
			}
		}
	case *jsast.ArrayLiteral:
		for i := range node.Value {
			if jsExpressionContainsNestedFunction(node.Value[i]) {
				return true
			}
		}
	case *jsast.ConditionalExpression:
		return jsExpressionContainsNestedFunction(node.Test) || jsExpressionContainsNestedFunction(node.Consequent) || jsExpressionContainsNestedFunction(node.Alternate)
	case *jsast.TemplateLiteral:
		for i := range node.Expressions {
			if jsExpressionContainsNestedFunction(node.Expressions[i]) {
				return true
			}
		}
	}
	return false
}

func jsStatementCapturesLoopNames(stmt jsast.Statement, names map[string]struct{}) bool {
	if stmt == nil || len(names) == 0 {
		return false
	}
	switch node := stmt.(type) {
	case *jsast.FunctionDeclaration:
		return jsFunctionCapturesLoopNames(node.Function, names)
	case *jsast.BlockStatement:
		for i := range node.List {
			if jsStatementCapturesLoopNames(node.List[i], names) {
				return true
			}
		}
	case *jsast.ExpressionStatement:
		return jsExpressionCapturesLoopNames(node.Expression, names)
	case *jsast.IfStatement:
		return jsExpressionCapturesLoopNames(node.Test, names) || jsStatementCapturesLoopNames(node.Consequent, names) || jsStatementCapturesLoopNames(node.Alternate, names)
	case *jsast.ForStatement:
		if node.Initializer != nil {
			switch init := node.Initializer.(type) {
			case *jsast.ForLoopInitializerExpression:
				if jsExpressionCapturesLoopNames(init.Expression, names) {
					return true
				}
			case *jsast.ForLoopInitializerVarDeclList:
				for _, b := range init.List {
					if b != nil && jsExpressionCapturesLoopNames(b.Initializer, names) {
						return true
					}
				}
			case *jsast.ForLoopInitializerLexicalDecl:
				for _, b := range init.LexicalDeclaration.List {
					if b != nil && jsExpressionCapturesLoopNames(b.Initializer, names) {
						return true
					}
				}
			}
		}
		return jsExpressionCapturesLoopNames(node.Test, names) || jsExpressionCapturesLoopNames(node.Update, names) || jsStatementCapturesLoopNames(node.Body, names)
	case *jsast.ForInStatement:
		return jsExpressionCapturesLoopNames(node.Source, names) || jsStatementCapturesLoopNames(node.Body, names)
	case *jsast.ForOfStatement:
		return jsExpressionCapturesLoopNames(node.Source, names) || jsStatementCapturesLoopNames(node.Body, names)
	case *jsast.ReturnStatement:
		return jsExpressionCapturesLoopNames(node.Argument, names)
	case *jsast.ThrowStatement:
		return jsExpressionCapturesLoopNames(node.Argument, names)
	case *jsast.WhileStatement:
		return jsExpressionCapturesLoopNames(node.Test, names) || jsStatementCapturesLoopNames(node.Body, names)
	case *jsast.DoWhileStatement:
		return jsExpressionCapturesLoopNames(node.Test, names) || jsStatementCapturesLoopNames(node.Body, names)
	case *jsast.SwitchStatement:
		if jsExpressionCapturesLoopNames(node.Discriminant, names) {
			return true
		}
		for i := range node.Body {
			clause := node.Body[i]
			if clause == nil {
				continue
			}
			if jsExpressionCapturesLoopNames(clause.Test, names) {
				return true
			}
			for j := range clause.Consequent {
				if jsStatementCapturesLoopNames(clause.Consequent[j], names) {
					return true
				}
			}
		}
	case *jsast.TryStatement:
		if jsStatementCapturesLoopNames(node.Body, names) {
			return true
		}
		if node.Catch != nil && jsStatementCapturesLoopNames(node.Catch.Body, names) {
			return true
		}
		if node.Finally != nil && jsStatementCapturesLoopNames(node.Finally, names) {
			return true
		}
	case *jsast.VariableStatement:
		for _, b := range node.List {
			if b != nil && jsExpressionCapturesLoopNames(b.Initializer, names) {
				return true
			}
		}
	case *jsast.LexicalDeclaration:
		for _, b := range node.List {
			if b != nil && jsExpressionCapturesLoopNames(b.Initializer, names) {
				return true
			}
		}
	case *jsast.UsingDeclaration:
		for _, b := range node.List {
			if b != nil && jsExpressionCapturesLoopNames(b.Initializer, names) {
				return true
			}
		}
	}
	return false
}

func jsExpressionCapturesLoopNames(expr jsast.Expression, names map[string]struct{}) bool {
	if expr == nil || len(names) == 0 {
		return false
	}
	switch node := expr.(type) {
	case *jsast.FunctionLiteral:
		return jsFunctionCapturesLoopNames(node, names)
	case *jsast.ArrowFunctionLiteral:
		return jsArrowFunctionCapturesLoopNames(node, names)
	case *jsast.ClassExpression:
		if node.Class == nil {
			return false
		}
		return jsClassCapturesLoopNames(node.Class, names)
	case *jsast.AssignExpression:
		return jsExpressionCapturesLoopNames(node.Left, names) || jsExpressionCapturesLoopNames(node.Right, names)
	case *jsast.BinaryExpression:
		return jsExpressionCapturesLoopNames(node.Left, names) || jsExpressionCapturesLoopNames(node.Right, names)
	case *jsast.UnaryExpression:
		return jsExpressionCapturesLoopNames(node.Operand, names)
	case *jsast.DotExpression:
		return jsExpressionCapturesLoopNames(node.Left, names)
	case *jsast.BracketExpression:
		return jsExpressionCapturesLoopNames(node.Left, names) || jsExpressionCapturesLoopNames(node.Member, names)
	case *jsast.CallExpression:
		if jsExpressionCapturesLoopNames(node.Callee, names) {
			return true
		}
		for i := range node.ArgumentList {
			if jsExpressionCapturesLoopNames(node.ArgumentList[i], names) {
				return true
			}
		}
	case *jsast.NewExpression:
		if jsExpressionCapturesLoopNames(node.Callee, names) {
			return true
		}
		for i := range node.ArgumentList {
			if jsExpressionCapturesLoopNames(node.ArgumentList[i], names) {
				return true
			}
		}
	case *jsast.ObjectLiteral:
		for i := range node.Value {
			switch p := node.Value[i].(type) {
			case *jsast.PropertyShort:
				if p.Initializer != nil {
					if jsExpressionCapturesLoopNames(p.Initializer, names) {
						return true
					}
				}
			case *jsast.PropertyKeyed:
				if jsExpressionCapturesLoopNames(p.Key, names) || jsExpressionCapturesLoopNames(p.Value, names) {
					return true
				}
			}
		}
	case *jsast.ArrayLiteral:
		for i := range node.Value {
			if jsExpressionCapturesLoopNames(node.Value[i], names) {
				return true
			}
		}
	case *jsast.ConditionalExpression:
		return jsExpressionCapturesLoopNames(node.Test, names) || jsExpressionCapturesLoopNames(node.Consequent, names) || jsExpressionCapturesLoopNames(node.Alternate, names)
	case *jsast.TemplateLiteral:
		for i := range node.Expressions {
			if jsExpressionCapturesLoopNames(node.Expressions[i], names) {
				return true
			}
		}
	case *jsast.SequenceExpression:
		for i := range node.Sequence {
			if jsExpressionCapturesLoopNames(node.Sequence[i], names) {
				return true
			}
		}
	}
	return false
}

func jsFunctionCapturesLoopNames(fn *jsast.FunctionLiteral, names map[string]struct{}) bool {
	if fn == nil || fn.Body == nil || len(names) == 0 {
		return false
	}
	visible := jsVisibleCaptureNamesForFunction(names, fn.Name, fn.ParameterList)
	if len(visible) == 0 {
		return false
	}
	for i := range fn.Body.List {
		if jsStatementReferencesLoopNames(fn.Body.List[i], visible) {
			return true
		}
	}
	return false
}

func jsArrowFunctionCapturesLoopNames(fn *jsast.ArrowFunctionLiteral, names map[string]struct{}) bool {
	if fn == nil || len(names) == 0 {
		return false
	}
	visible := jsVisibleCaptureNamesForFunction(names, nil, fn.ParameterList)
	if len(visible) == 0 {
		return false
	}
	switch body := fn.Body.(type) {
	case *jsast.ExpressionBody:
		return jsExpressionReferencesLoopNames(body.Expression, visible)
	case *jsast.BlockStatement:
		for i := range body.List {
			if jsStatementReferencesLoopNames(body.List[i], visible) {
				return true
			}
		}
	}
	return false
}

func jsClassCapturesLoopNames(class *jsast.ClassLiteral, names map[string]struct{}) bool {
	if class == nil || len(names) == 0 {
		return false
	}
	if jsExpressionCapturesLoopNames(class.SuperClass, names) {
		return true
	}
	for i := range class.Body {
		switch el := class.Body[i].(type) {
		case *jsast.FieldDefinition:
			if jsExpressionCapturesLoopNames(el.Key, names) || jsExpressionCapturesLoopNames(el.Initializer, names) {
				return true
			}
		case *jsast.MethodDefinition:
			if jsExpressionCapturesLoopNames(el.Key, names) || jsFunctionCapturesLoopNames(el.Body, names) {
				return true
			}
		case *jsast.ClassStaticBlock:
			if el.Block != nil {
				for j := range el.Block.List {
					if jsStatementCapturesLoopNames(el.Block.List[j], names) {
						return true
					}
				}
			}
		}
	}
	return false
}

func jsVisibleCaptureNamesForFunction(names map[string]struct{}, fnName *jsast.Identifier, params *jsast.ParameterList) map[string]struct{} {
	if len(names) == 0 {
		return nil
	}
	visible := make(map[string]struct{}, len(names))
	for name := range names {
		visible[name] = struct{}{}
	}
	if fnName != nil {
		delete(visible, fnName.Name.String())
	}
	if params != nil {
		for _, b := range params.List {
			if b == nil || b.Target == nil {
				continue
			}
			paramNames := make([]string, 0, 2)
			jsExtractBindingNames(b.Target, &paramNames)
			for _, n := range paramNames {
				delete(visible, n)
			}
		}
		if params.Rest != nil {
			restNames := make([]string, 0, 2)
			if target, ok := params.Rest.(jsast.BindingTarget); ok {
				jsExtractBindingNames(target, &restNames)
			}
			for _, n := range restNames {
				delete(visible, n)
			}
		}
	}
	if len(visible) == 0 {
		return nil
	}
	return visible
}

func jsStatementReferencesLoopNames(stmt jsast.Statement, names map[string]struct{}) bool {
	if stmt == nil || len(names) == 0 {
		return false
	}
	switch node := stmt.(type) {
	case *jsast.FunctionDeclaration:
		return jsFunctionCapturesLoopNames(node.Function, names)
	case *jsast.BlockStatement:
		for i := range node.List {
			if jsStatementReferencesLoopNames(node.List[i], names) {
				return true
			}
		}
	case *jsast.ExpressionStatement:
		return jsExpressionReferencesLoopNames(node.Expression, names)
	case *jsast.IfStatement:
		return jsExpressionReferencesLoopNames(node.Test, names) || jsStatementReferencesLoopNames(node.Consequent, names) || jsStatementReferencesLoopNames(node.Alternate, names)
	case *jsast.ForStatement:
		if node.Initializer != nil {
			switch init := node.Initializer.(type) {
			case *jsast.ForLoopInitializerExpression:
				if jsExpressionReferencesLoopNames(init.Expression, names) {
					return true
				}
			case *jsast.ForLoopInitializerVarDeclList:
				for _, b := range init.List {
					if b != nil && jsExpressionReferencesLoopNames(b.Initializer, names) {
						return true
					}
				}
			case *jsast.ForLoopInitializerLexicalDecl:
				for _, b := range init.LexicalDeclaration.List {
					if b != nil && jsExpressionReferencesLoopNames(b.Initializer, names) {
						return true
					}
				}
			}
		}
		return jsExpressionReferencesLoopNames(node.Test, names) || jsExpressionReferencesLoopNames(node.Update, names) || jsStatementReferencesLoopNames(node.Body, names)
	case *jsast.ForInStatement:
		return jsExpressionReferencesLoopNames(node.Source, names) || jsStatementReferencesLoopNames(node.Body, names)
	case *jsast.ForOfStatement:
		return jsExpressionReferencesLoopNames(node.Source, names) || jsStatementReferencesLoopNames(node.Body, names)
	case *jsast.ReturnStatement:
		return jsExpressionReferencesLoopNames(node.Argument, names)
	case *jsast.ThrowStatement:
		return jsExpressionReferencesLoopNames(node.Argument, names)
	case *jsast.WhileStatement:
		return jsExpressionReferencesLoopNames(node.Test, names) || jsStatementReferencesLoopNames(node.Body, names)
	case *jsast.DoWhileStatement:
		return jsExpressionReferencesLoopNames(node.Test, names) || jsStatementReferencesLoopNames(node.Body, names)
	case *jsast.SwitchStatement:
		if jsExpressionReferencesLoopNames(node.Discriminant, names) {
			return true
		}
		for i := range node.Body {
			clause := node.Body[i]
			if clause == nil {
				continue
			}
			if jsExpressionReferencesLoopNames(clause.Test, names) {
				return true
			}
			for j := range clause.Consequent {
				if jsStatementReferencesLoopNames(clause.Consequent[j], names) {
					return true
				}
			}
		}
	case *jsast.TryStatement:
		if jsStatementReferencesLoopNames(node.Body, names) {
			return true
		}
		if node.Catch != nil && jsStatementReferencesLoopNames(node.Catch.Body, names) {
			return true
		}
		if node.Finally != nil && jsStatementReferencesLoopNames(node.Finally, names) {
			return true
		}
	case *jsast.VariableStatement:
		for _, b := range node.List {
			if b != nil && jsExpressionReferencesLoopNames(b.Initializer, names) {
				return true
			}
		}
	case *jsast.LexicalDeclaration:
		for _, b := range node.List {
			if b != nil && jsExpressionReferencesLoopNames(b.Initializer, names) {
				return true
			}
		}
	case *jsast.UsingDeclaration:
		for _, b := range node.List {
			if b != nil && jsExpressionReferencesLoopNames(b.Initializer, names) {
				return true
			}
		}
	}
	return false
}

func jsExpressionReferencesLoopNames(expr jsast.Expression, names map[string]struct{}) bool {
	if expr == nil || len(names) == 0 {
		return false
	}
	switch node := expr.(type) {
	case *jsast.Identifier:
		_, ok := names[node.Name.String()]
		return ok
	case *jsast.FunctionLiteral:
		return jsFunctionCapturesLoopNames(node, names)
	case *jsast.ArrowFunctionLiteral:
		return jsArrowFunctionCapturesLoopNames(node, names)
	case *jsast.ClassExpression:
		if node.Class == nil {
			return false
		}
		return jsClassCapturesLoopNames(node.Class, names)
	case *jsast.AssignExpression:
		return jsExpressionReferencesLoopNames(node.Left, names) || jsExpressionReferencesLoopNames(node.Right, names)
	case *jsast.BinaryExpression:
		return jsExpressionReferencesLoopNames(node.Left, names) || jsExpressionReferencesLoopNames(node.Right, names)
	case *jsast.UnaryExpression:
		return jsExpressionReferencesLoopNames(node.Operand, names)
	case *jsast.DotExpression:
		return jsExpressionReferencesLoopNames(node.Left, names)
	case *jsast.BracketExpression:
		return jsExpressionReferencesLoopNames(node.Left, names) || jsExpressionReferencesLoopNames(node.Member, names)
	case *jsast.CallExpression:
		if jsExpressionReferencesLoopNames(node.Callee, names) {
			return true
		}
		for i := range node.ArgumentList {
			if jsExpressionReferencesLoopNames(node.ArgumentList[i], names) {
				return true
			}
		}
	case *jsast.NewExpression:
		if jsExpressionReferencesLoopNames(node.Callee, names) {
			return true
		}
		for i := range node.ArgumentList {
			if jsExpressionReferencesLoopNames(node.ArgumentList[i], names) {
				return true
			}
		}
	case *jsast.ObjectLiteral:
		for i := range node.Value {
			switch p := node.Value[i].(type) {
			case *jsast.PropertyShort:
				if p.Initializer != nil {
					if jsExpressionReferencesLoopNames(p.Initializer, names) {
						return true
					}
				} else {
					if _, ok := names[p.Name.Name.String()]; ok {
						return true
					}
				}
			case *jsast.PropertyKeyed:
				if jsExpressionReferencesLoopNames(p.Key, names) || jsExpressionReferencesLoopNames(p.Value, names) {
					return true
				}
			}
		}
	case *jsast.ArrayLiteral:
		for i := range node.Value {
			if jsExpressionReferencesLoopNames(node.Value[i], names) {
				return true
			}
		}
	case *jsast.ConditionalExpression:
		return jsExpressionReferencesLoopNames(node.Test, names) || jsExpressionReferencesLoopNames(node.Consequent, names) || jsExpressionReferencesLoopNames(node.Alternate, names)
	case *jsast.TemplateLiteral:
		for i := range node.Expressions {
			if jsExpressionReferencesLoopNames(node.Expressions[i], names) {
				return true
			}
		}
	case *jsast.SequenceExpression:
		for i := range node.Sequence {
			if jsExpressionReferencesLoopNames(node.Sequence[i], names) {
				return true
			}
		}
	}
	return false
}

func (c *Compiler) compileJScriptFunctionLiteral(fn *jsast.FunctionLiteral, fallbackName string, isClassConstructor bool) {
	jumpOverBody := c.emitJSJump(OpJSJump)
	bodyStart := len(c.bytecode)

	prevLocalEnabled := c.jsLocalEnabled
	prevLocalSlotCount := c.jsLocalSlotCount
	prevLocalScopeStack := c.jsLocalScopeStack
	prevGenerator := c.jsInGeneratorFunction
	c.jsInGeneratorFunction = fn != nil && fn.Generator
	defer func() { c.jsInGeneratorFunction = prevGenerator }()
	canUseLocalSlots := !jsFunctionContainsNestedFunction(fn) && (fn == nil || !fn.Generator)
	c.jsLocalEnabled = canUseLocalSlots
	c.jsLocalSlotCount = 0
	c.jsLocalScopeStack = make([]jsLocalScope, 0, 8)
	if c.jsLocalEnabled {
		c.jsPushLocalScope(true)
	}

	if fn != nil && fn.ParameterList != nil {
		for _, b := range fn.ParameterList.List {
			if b == nil || b.Target == nil {
				continue
			}
			if p, ok := b.Target.(*jsast.Identifier); ok {
				if c.jsLocalEnabled {
					slot := c.jsDeclareFunctionLocal(p.Name.String())
					if slot >= 0 {
						nameIdx := c.addConstant(NewString(p.Name.String()))
						c.emit(OpJSGetName, nameIdx)
						c.emit(OpJSSetLocal, slot)
					}
				}
			}
		}
		if fn.ParameterList.Rest != nil {
			if restID, ok := fn.ParameterList.Rest.(*jsast.Identifier); ok && c.jsLocalEnabled {
				slot := c.jsDeclareFunctionLocal(restID.Name.String())
				if slot >= 0 {
					nameIdx := c.addConstant(NewString(restID.Name.String()))
					c.emit(OpJSGetName, nameIdx)
					c.emit(OpJSSetLocal, slot)
				}
			}
		}
	}

	// Emit default parameter guards before the function body.
	c.compileJScriptDefaultParamGuards(fn.ParameterList)

	// Inject field initialization for base class constructors
	if isClassConstructor && !c.jsIsDerivedConstructor {
		c.compileJScriptClassFields()
	}

	// Detect "use strict" directive at the beginning of the function body
	hasStrictMode, _ := c.detectUseStrictDirective(fn.Body.List)
	prevStrictMode := c.jsStrictMode
	if hasStrictMode {
		c.jsStrictMode = true
	}

	if fn.Body != nil {
		for i := range fn.Body.List {
			c.compileJScriptStatement(fn.Body.List[i])
		}
	}
	c.emit(OpJSLoadUndefined)
	c.emit(OpJSReturn)
	bodyEnd := len(c.bytecode)
	localCount := c.jsLocalSlotCount
	if c.jsLocalEnabled {
		c.jsPopLocalScope()
	}
	c.jsLocalEnabled = prevLocalEnabled
	c.jsLocalSlotCount = prevLocalSlotCount
	c.jsLocalScopeStack = prevLocalScopeStack
	c.patchJSJump(jumpOverBody)

	name := fallbackName
	if fn.Name != nil {
		name = fn.Name.Name.String()
	}
	params := make([]string, 0)
	if fn.ParameterList != nil {
		params = make([]string, 0, len(fn.ParameterList.List))
		paramSeen := make(map[string]struct{}, len(fn.ParameterList.List))
		for _, b := range fn.ParameterList.List {
			if b == nil || b.Target == nil {
				continue
			}
			if p, ok := b.Target.(*jsast.Identifier); ok {
				paramName := p.Name.String()
				// In strict mode, eval/arguments as parameter names are SyntaxErrors
				if c.jsStrictMode && jsIsRestrictedIdentifier(paramName) {
					jsErr := jscript.NewJSSyntaxError(jscript.IllegalAssignment, 0, 0)
					jsErr.WithASPDescription(fmt.Sprintf("parameter name '%s' is not allowed in strict mode", paramName))
					if c.sourceName != "" {
						jsErr.WithFile(c.sourceName)
					}
					panic(jsErr)
				}
				// In strict mode, duplicate parameter names are SyntaxErrors
				if c.jsStrictMode {
					if _, duplicate := paramSeen[paramName]; duplicate {
						jsErr := jscript.NewJSSyntaxError(jscript.SyntaxError, 0, 0)
						jsErr.WithASPDescription(fmt.Sprintf("duplicate parameter name '%s' not allowed in strict mode", paramName))
						if c.sourceName != "" {
							jsErr.WithFile(c.sourceName)
						}
						panic(jsErr)
					}
					paramSeen[paramName] = struct{}{}
				}
				params = append(params, paramName)
			}
		}
		if fn.ParameterList.Rest != nil {
			if restID, ok := fn.ParameterList.Rest.(*jsast.Identifier); ok {
				restName := restID.Name.String()
				if c.jsStrictMode && jsIsRestrictedIdentifier(restName) {
					jsErr := jscript.NewJSSyntaxError(jscript.IllegalAssignment, 0, 0)
					jsErr.WithASPDescription(fmt.Sprintf("parameter name '%s' is not allowed in strict mode", restName))
					if c.sourceName != "" {
						jsErr.WithFile(c.sourceName)
					}
					panic(jsErr)
				}
				if c.jsStrictMode {
					if _, duplicate := paramSeen[restName]; duplicate {
						jsErr := jscript.NewJSSyntaxError(jscript.SyntaxError, 0, 0)
						jsErr.WithASPDescription(fmt.Sprintf("duplicate parameter name '%s' not allowed in strict mode", restName))
						if c.sourceName != "" {
							jsErr.WithFile(c.sourceName)
						}
						panic(jsErr)
					}
				}
				params = append(params, jsRestParamTemplatePrefix+restName)
			}
		}
	}

	if isClassConstructor {
		params = append(params, jsClassConstructorFlag)
	}
	if c.jsIsDerivedConstructor {
		params = append(params, jsDerivedConstructorFlag)
	}
	if c.jsStrictMode {
		params = append(params, jsStrictModeFlag)
	}
	// Restore strict mode after we've used it to set the flag in params
	c.jsStrictMode = prevStrictMode

	if fn.Generator {
		params = append(params, jsGeneratorFlag)
	}
	if fn.Async {
		params = append(params, jsAsyncFlag)
	}
	if localCount > 0 {
		params = append(params, "__js_local_count__:"+strconv.Itoa(localCount))
	}

	templateIdx := c.addConstant(Value{
		Type:  VTJSFunctionTemplate,
		Num:   int64(bodyStart),
		Flt:   float64(bodyEnd),
		Str:   name,
		Names: params,
	})
	c.emit(OpJSCreateClosure, templateIdx)
}

func (c *Compiler) compileJScriptUpdateExpression(node *jsast.UnaryExpression) bool {
	switch operand := node.Operand.(type) {
	case *jsast.Identifier:
		name := operand.Name.String()
		nameIdx := c.addConstant(NewString(name))
		slot, isLocal := c.jsResolveLocalSlot(name)
		isInt := c.jsGetLocalType(name) == jsTypeInteger

		switch node.Operator {
		case jstoken.INCREMENT:
			if isLocal {
				if node.Postfix {
					if !c.jsInGeneratorFunction {
						c.emit(OpJSGetLocal, slot)
						c.emit(OpJSIncLocal, slot)
						c.emit(OpJSPop)
						return true
					}
				} else {
					c.emit(OpJSIncLocal, slot)
					return true
				}
			}
			if isInt && !node.Postfix {
				c.emit(OpJSIncInt, nameIdx)
				return true
			}
			if node.Postfix {
				c.emit(OpJSPostIncrement, nameIdx)
			} else {
				c.emit(OpJSPreIncrement, nameIdx)
			}
			return true
		case jstoken.DECREMENT:
			if isLocal {
				if node.Postfix {
					c.emit(OpJSGetLocal, slot)
					c.emit(OpJSDecLocal, slot)
					c.emit(OpJSPop)
				} else {
					c.emit(OpJSDecLocal, slot)
				}
				return true
			}
			if node.Postfix {
				c.emit(OpJSPostDecrement, nameIdx)
			} else {
				c.emit(OpJSPreDecrement, nameIdx)
			}
			return true
		}
	case *jsast.PrivateDotExpression:
		c.compileJScriptExpression(operand.Left)
		nameIdx := c.addConstant(NewString("\x00__priv_" + operand.Identifier.Name.String()))
		switch node.Operator {
		case jstoken.INCREMENT:
			if node.Postfix {
				c.emit(OpJSPostMemberIncrement, nameIdx)
			} else {
				c.emit(OpJSPreMemberIncrement, nameIdx)
			}
			return true
		case jstoken.DECREMENT:
			if node.Postfix {
				c.emit(OpJSPostMemberDecrement, nameIdx)
			} else {
				c.emit(OpJSPreMemberDecrement, nameIdx)
			}
			return true
		}
	case *jsast.DotExpression:
		c.compileJScriptExpression(operand.Left)
		nameIdx := c.addConstant(NewString(operand.Identifier.Name.String()))
		switch node.Operator {
		case jstoken.INCREMENT:
			if node.Postfix {
				c.emit(OpJSPostMemberIncrement, nameIdx)
			} else {
				c.emit(OpJSPreMemberIncrement, nameIdx)
			}
			return true
		case jstoken.DECREMENT:
			if node.Postfix {
				c.emit(OpJSPostMemberDecrement, nameIdx)
			} else {
				c.emit(OpJSPreMemberDecrement, nameIdx)
			}
			return true
		}
	case *jsast.BracketExpression:
		c.compileJScriptExpression(operand.Left)
		c.compileJScriptExpression(operand.Member)
		switch node.Operator {
		case jstoken.INCREMENT:
			if node.Postfix {
				c.emit(OpJSPostIndexIncrement)
			} else {
				c.emit(OpJSPreIndexIncrement)
			}
			return true
		case jstoken.DECREMENT:
			if node.Postfix {
				c.emit(OpJSPostIndexDecrement)
			} else {
				c.emit(OpJSPreIndexDecrement)
			}
			return true
		}
	}
	return false
}

func jsExtractBindingNames(target jsast.BindingTarget, names *[]string) {
	if target == nil {
		return
	}
	switch t := target.(type) {
	case *jsast.Identifier:
		*names = append(*names, t.Name.String())
	case *jsast.ObjectPattern:
		for _, prop := range t.Properties {
			switch p := prop.(type) {
			case *jsast.PropertyShort:
				jsExtractBindingNames(&p.Name, names)
			case *jsast.PropertyKeyed:
				if bt, ok := p.Value.(jsast.BindingTarget); ok {
					jsExtractBindingNames(bt, names)
				}
			}
		}
		if t.Rest != nil {
			if bt, ok := t.Rest.(jsast.BindingTarget); ok {
				jsExtractBindingNames(bt, names)
			}
		}
	case *jsast.ArrayPattern:
		for _, elt := range t.Elements {
			if elt != nil {
				if bt, ok := elt.(jsast.BindingTarget); ok {
					jsExtractBindingNames(bt, names)
				}
			}
		}
		if t.Rest != nil {
			if bt, ok := t.Rest.(jsast.BindingTarget); ok {
				jsExtractBindingNames(bt, names)
			}
		}
	}
}

func jsBindingIdentifierName(target jsast.BindingTarget) (string, bool) {
	if id, ok := target.(*jsast.Identifier); ok {
		return id.Name.String(), true
	}
	return "", false
}

// jsIsRestrictedIdentifier returns true if name is "eval" or "arguments",
// which are restricted in strict mode.
func jsIsRestrictedIdentifier(name string) bool {
	return strings.EqualFold(name, "eval") || strings.EqualFold(name, "arguments")
}

// jsGetBlockLexicalNames returns lists of names for let and const declarations in the block.
func jsGetBlockLexicalNames(stmts []jsast.Statement) ([]string, []string) {
	var letNames []string
	var constNames []string
	for _, s := range stmts {
		if decl, ok := s.(*jsast.LexicalDeclaration); ok {
			for _, binding := range decl.List {
				if decl.Token == jstoken.CONST {
					jsExtractBindingNames(binding.Target, &constNames)
				} else {
					jsExtractBindingNames(binding.Target, &letNames)
				}
			}
		} else if decl, ok := s.(*jsast.UsingDeclaration); ok {
			for _, binding := range decl.List {
				jsExtractBindingNames(binding.Target, &letNames)
			}
		} else if decl, ok := s.(*jsast.ClassDeclaration); ok {
			if decl.Class != nil && decl.Class.Name != nil {
				letNames = append(letNames, decl.Class.Name.Name.String())
			}
		}
	}
	return letNames, constNames
}

func (c *Compiler) compileJScriptDestructuring(target jsast.Expression, isConst bool, isLet bool, isVar bool) {
	if target == nil {
		return
	}
	switch t := target.(type) {
	case *jsast.Identifier:
		name := t.Name.String()
		if c.jsStrictMode && jsIsRestrictedIdentifier(name) {
			jsErr := jscript.NewJSSyntaxError(jscript.IllegalAssignment, 0, 0)
			jsErr.WithASPDescription(fmt.Sprintf("cannot use '%s' as a variable name in strict mode", name))
			if c.sourceName != "" {
				jsErr.WithFile(c.sourceName)
			}
			panic(jsErr)
		}
		nameIdx := c.addConstant(NewString(name))
		if isVar {
			if c.jsLocalEnabled {
				slot := c.jsDeclareFunctionLocal(name)
				if slot >= 0 {
					c.emit(OpJSSetLocal, slot)
					break
				}
			}
			c.emit(OpJSDeclareName, nameIdx)
			c.emit(OpJSSetName, nameIdx)
		} else if isConst {
			if c.jsLocalEnabled {
				c.jsAddLocalBarrier(name)
			}
			c.emitConstInitialize(nameIdx)
		} else if isLet {
			if c.jsLocalEnabled {
				c.jsAddLocalBarrier(name)
			}
			c.emit(OpJSLetDeclare, nameIdx)
			c.emit(OpJSSetName, nameIdx)
		} else {
			// Normal assignment
			if slot, ok := c.jsResolveLocalSlot(name); ok {
				c.emit(OpJSSetLocal, slot)
			} else {
				c.emit(OpJSSetName, nameIdx)
			}
		}
	case *jsast.AssignExpression:
		if t.Operator == jstoken.ASSIGN {
			jump := c.emitJSJump(OpJSJumpIfNotUndefined)
			c.emit(OpJSPop)
			c.compileJScriptExpression(t.Right)
			c.patchJSJump(jump)
			c.compileJScriptDestructuring(t.Left, isConst, isLet, isVar)
		} else {
			c.emit(OpJSPop)
		}
	case *jsast.ObjectPattern:
		c.emit(OpJSRequireObject)
		var excludedStatic []string
		for _, prop := range t.Properties {
			switch p := prop.(type) {
			case *jsast.PropertyShort:
				name := p.Name.Name.String()
				excludedStatic = append(excludedStatic, name)
				nameIdx := c.addConstant(NewString(name))
				c.emit(OpJSDup)
				c.emitJSMemberGet(nameIdx)
				if p.Initializer != nil {
					jump := c.emitJSJump(OpJSJumpIfNotUndefined)
					c.emit(OpJSPop)
					c.compileJScriptExpression(p.Initializer)
					c.patchJSJump(jump)
				}
				c.compileJScriptDestructuring(&p.Name, isConst, isLet, isVar)
			case *jsast.PropertyKeyed:
				c.emit(OpJSDup)
				if p.Computed {
					c.compileJScriptExpression(p.Key)
					c.emit(OpJSIndexGet)
				} else {
					key := ""
					if id, ok := p.Key.(*jsast.Identifier); ok {
						key = id.Name.String()
					} else if lit, ok := p.Key.(*jsast.StringLiteral); ok {
						key = lit.Value.String()
					}
					excludedStatic = append(excludedStatic, key)
					c.emitJSMemberGet(c.addConstant(NewString(key)))
				}
				c.compileJScriptDestructuring(p.Value, isConst, isLet, isVar)
			}
		}
		if t.Rest != nil {
			c.emit(OpJSDup)
			c.bytecode = append(c.bytecode, byte(OpJSObjectRest))
			c.bytecode = append(c.bytecode, byte(len(excludedStatic)>>8), byte(len(excludedStatic)&0xFF))
			for _, key := range excludedStatic {
				idx := c.addConstant(NewString(key))
				c.bytecode = append(c.bytecode, byte(idx>>8), byte(idx&0xFF))
			}
			c.bytecode = append(c.bytecode, 0, 0) // 0 dynamic exclusions
			c.compileJScriptDestructuring(t.Rest, isConst, isLet, isVar)
		}
		c.emit(OpJSPop) // Pop the source object
	case *jsast.ArrayPattern:
		c.emit(OpJSGetIterator)
		for _, elt := range t.Elements {
			c.emit(OpJSIteratorNext)
			if elt != nil {
				c.compileJScriptDestructuring(elt, isConst, isLet, isVar)
			} else {
				// Elision: [,,]
				c.emit(OpJSPop)
			}
		}
		if t.Rest != nil {
			c.emit(OpJSDup)
			c.emit(OpJSCollectRest)
			c.compileJScriptDestructuring(t.Rest, isConst, isLet, isVar)
		}
		c.emit(OpJSPop) // Pop the iterator
	default:
		// Unsupported pattern element or regular expression (e.g. [a.b])
		// For now we just pop.
		c.emit(OpJSPop)
	}
}

// compileJScriptLexicalDeclaration emits block-scoped let/const declarations.
func (c *Compiler) compileJScriptLexicalDeclaration(node *jsast.LexicalDeclaration) {
	isConst := node.Token == jstoken.CONST
	for _, binding := range node.List {
		if isConst {
			if binding.Initializer != nil {
				c.compileJScriptExpression(binding.Initializer)
				t := jsTypeUnknown
				if c.jsInferredType(binding.Initializer) == jsTypeInteger {
					t = jsTypeInteger
				}
				c.compileJScriptDestructuring(binding.Target, true, false, false)
				if t == jsTypeInteger {
					if id, ok := binding.Target.(*jsast.Identifier); ok {
						c.jsSetLocalType(id.Name.String(), jsTypeInteger)
					}
				}
			} else {
				// const without initializer is a SyntaxError per spec
				jsErr := jscript.NewJSSyntaxError(jscript.SyntaxError, 0, 0)
				jsErr.WithASPDescription("missing initializer in const declaration")
				if c.sourceName != "" {
					jsErr.WithFile(c.sourceName)
				}
				panic(jsErr)
			}
		} else {
			// let
			if binding.Initializer != nil {
				c.compileJScriptExpression(binding.Initializer)
				t := jsTypeUnknown
				if c.jsInferredType(binding.Initializer) == jsTypeInteger {
					t = jsTypeInteger
				}
				c.compileJScriptDestructuring(binding.Target, false, true, false)
				if t == jsTypeInteger {
					if id, ok := binding.Target.(*jsast.Identifier); ok {
						c.jsSetLocalType(id.Name.String(), jsTypeInteger)
					}
				}
			} else {
				// let x; -> declare x
				if id, ok := binding.Target.(*jsast.Identifier); ok {
					if c.jsLocalEnabled {
						c.jsAddLocalBarrier(id.Name.String())
					}
					nameIdx := c.addConstant(NewString(id.Name.String()))
					c.emit(OpJSLetDeclare, nameIdx)
				}
			}
		}
	}
}

// emitConstInitialize emits the OpJSConstInitialize opcode for the given name index.
func (c *Compiler) emitConstInitialize(nameIdx int) {
	c.bytecode = append(c.bytecode, byte(OpJSConstInitialize))
	c.bytecode = append(c.bytecode, byte(nameIdx>>8), byte(nameIdx&0xFF))
}

func (c *Compiler) emitJSJump(op OpCode) int {
	pos := c.emit(op, 0)
	return pos + 1
}

func (c *Compiler) patchJSJump(offsetIndex int) {
	c.patchJSJumpTo(offsetIndex, len(c.bytecode))
}

func (c *Compiler) patchJSJumpTo(offsetIndex int, jumpTarget int) {
	if offsetIndex < 0 || offsetIndex+4 > len(c.bytecode) {
		panic("js jump patch out of range")
	}
	c.bytecode[offsetIndex] = byte((jumpTarget >> 24) & 0xFF)
	c.bytecode[offsetIndex+1] = byte((jumpTarget >> 16) & 0xFF)
	c.bytecode[offsetIndex+2] = byte((jumpTarget >> 8) & 0xFF)
	c.bytecode[offsetIndex+3] = byte(jumpTarget & 0xFF)
}

// ---------------------------------------------------------------------------
// JScript compile-time constant folding (AST pre-pass)
// ---------------------------------------------------------------------------

// compileJScriptTemplateLiteral compiles an ES6 template literal into a series of
// string concatenation operations on the stack, producing a single string value.
// Multi-line strings are supported natively since newlines are preserved in the
// element text. Tagged templates are not yet supported and emit undefined.
func (c *Compiler) compileJScriptTemplateLiteral(node *jsast.TemplateLiteral) {
	// Tagged templates: not supported in this initial implementation.
	if node.Tag != nil {
		c.emit(OpJSLoadUndefined)
		return
	}
	// Plain template with no expressions: emit as a string constant.
	if len(node.Expressions) == 0 {
		str := ""
		if len(node.Elements) > 0 && node.Elements[0].Valid {
			str = node.Elements[0].Parsed.String()
		}
		c.emit(OpConstant, c.addConstant(NewString(str)))
		return
	}
	// Build: elem[0] + expr[0] + elem[1] + expr[1] + ... + elem[n].
	// Starting with the first static element ensures the concatenation chain
	// begins as a string, triggering JS string coercion for all following values.
	firstElem := ""
	if len(node.Elements) > 0 && node.Elements[0].Valid {
		firstElem = node.Elements[0].Parsed.String()
	}
	c.emit(OpConstant, c.addConstant(NewString(firstElem)))
	for i, exprNode := range node.Expressions {
		c.compileJScriptExpression(exprNode)
		c.emit(OpJSAdd)
		elemStr := ""
		if i+1 < len(node.Elements) && node.Elements[i+1].Valid {
			elemStr = node.Elements[i+1].Parsed.String()
		}
		c.emit(OpConstant, c.addConstant(NewString(elemStr)))
		c.emit(OpJSAdd)
	}
}

// compileJScriptArrowFunctionLiteral compiles an ES6 arrow function expression.
// Arrow functions capture 'this' lexically from the enclosing scope; they do not
// bind their own 'this' when called. Concise bodies (x => expr) emit an implicit
// return; block bodies behave like regular functions.
func (c *Compiler) compileJScriptArrowFunctionLiteral(fn *jsast.ArrowFunctionLiteral) {
	jumpOverBody := c.emitJSJump(OpJSJump)
	bodyStart := len(c.bytecode)

	// Emit default parameter guards for any parameters that have an initializer.
	c.compileJScriptDefaultParamGuards(fn.ParameterList)

	switch body := fn.Body.(type) {
	case *jsast.ExpressionBody:
		// Concise body: `(x) => x * 2` — expression result is implicitly returned.
		c.compileJScriptExpression(body.Expression)
		c.emit(OpJSReturn)
	case *jsast.BlockStatement:
		for i := range body.List {
			c.compileJScriptStatement(body.List[i])
		}
		c.emit(OpJSLoadUndefined)
		c.emit(OpJSReturn)
	default:
		c.emit(OpJSLoadUndefined)
		c.emit(OpJSReturn)
	}

	bodyEnd := len(c.bytecode)
	c.patchJSJump(jumpOverBody)

	params := make([]string, 0)
	if fn.ParameterList != nil {
		params = make([]string, 0, len(fn.ParameterList.List))
		for _, b := range fn.ParameterList.List {
			if b == nil || b.Target == nil {
				continue
			}
			if p, ok := b.Target.(*jsast.Identifier); ok {
				params = append(params, p.Name.String())
			}
		}
		if fn.ParameterList.Rest != nil {
			if restID, ok := fn.ParameterList.Rest.(*jsast.Identifier); ok {
				params = append(params, jsRestParamTemplatePrefix+restID.Name.String())
			}
		}
	}

	templateIdx := c.addConstant(Value{
		Type:  VTJSArrowFunctionTemplate,
		Num:   int64(bodyStart),
		Flt:   float64(bodyEnd),
		Str:   "",
		Names: params,
	})
	c.emit(OpJSCreateClosure, templateIdx)
}

func (c *Compiler) compileJScriptClassDeclaration(node *jsast.ClassDeclaration) {
	if node.Class == nil {
		return
	}

	name := ""
	if node.Class.Name != nil {
		name = node.Class.Name.Name.String()
	}
	if name == "" {
		// Class declarations must have a name
		jsErr := jscript.NewJSSyntaxError(jscript.SyntaxError, 0, 0)
		jsErr.WithASPDescription("class declarations must have a name")
		if c.sourceName != "" {
			jsErr.WithFile(c.sourceName)
		}
		panic(jsErr)
	}

	nameIdx := c.addConstant(NewString(name))
	// Classes are block-scoped and NOT hoisted. Treat them like 'let'.
	c.emit(OpJSLetDeclare, nameIdx)

	// Compile class literal (which produces the constructor function)
	c.compileJScriptClassLiteral(node.Class)

	// Initialize the binding
	c.emit(OpJSSetName, nameIdx)
}

func (c *Compiler) compileJScriptClassLiteral(node *jsast.ClassLiteral) {
	// ES6 classes implicitly run in Strict Mode
	oldStrict := c.jsStrictMode
	c.jsStrictMode = true
	defer func() { c.jsStrictMode = oldStrict }()

	hasSuperClass := node.SuperClass != nil
	isNullHeritage := false
	if hasSuperClass {
		if _, ok := node.SuperClass.(*jsast.NullLiteral); ok {
			isNullHeritage = true
		}
	}
	isDerived := hasSuperClass && !isNullHeritage
	if hasSuperClass {
		// Evaluate superclass expression
		c.compileJScriptExpression(node.SuperClass)
	}

	// Collect instance fields to be initialized in the constructor
	oldFields := c.jsClassFields
	c.jsClassFields = nil
	for _, el := range node.Body {
		if field, ok := el.(*jsast.FieldDefinition); ok && !field.Static {
			c.jsClassFields = append(c.jsClassFields, field)
		}
	}
	defer func() { c.jsClassFields = oldFields }()

	var ctor *jsast.MethodDefinition
	for _, el := range node.Body {
		if md, ok := el.(*jsast.MethodDefinition); ok && md.Kind == jsast.PropertyKindConstructor {
			if ctor != nil {
				// This should have been caught by the parser, but we re-check for safety
				jsErr := jscript.NewJSSyntaxError(jscript.SyntaxError, 0, 0)
				jsErr.WithASPDescription("A class may only have one constructor")
				if c.sourceName != "" {
					jsErr.WithFile(c.sourceName)
				}
				panic(jsErr)
			}
			ctor = md
		}
	}

	className := ""
	if node.Name != nil {
		className = node.Name.Name.String()
	}

	// Compile the constructor
	oldIsDerived := c.jsIsDerivedConstructor
	c.jsIsDerivedConstructor = isDerived
	defer func() { c.jsIsDerivedConstructor = oldIsDerived }()

	if ctor == nil {
		// Create default constructor
		fn := &jsast.FunctionLiteral{
			ParameterList: &jsast.ParameterList{},
			Body:          &jsast.BlockStatement{},
		}
		if isDerived {
			// default derived constructor: constructor(...args) { super(...args); }
			// Simplified for now: just call super()
			fn.Body.List = []jsast.Statement{
				&jsast.ExpressionStatement{
					Expression: &jsast.CallExpression{
						Callee:       &jsast.SuperExpression{},
						ArgumentList: nil,
					},
				},
			}
		}
		c.compileJScriptFunctionLiteral(fn, className, true)
	} else {
		// We reuse compileJScriptFunctionLiteral but with a flag if it's a constructor.
		c.compileJScriptFunctionLiteral(ctor.Body, className, true)
	}

	// Constructor function is now on top of the stack.
	// 1. Static Inheritance: if derived, wire constructor.__proto__ = SuperClass
	if hasSuperClass {
		c.emit(OpJSClassInherit)
	}

	for _, el := range node.Body {
		md, ok := el.(*jsast.MethodDefinition)
		if !ok || md.Kind == jsast.PropertyKindConstructor {
			continue
		}

		c.emit(OpJSDup) // Duplicate the constructor

		if !md.Static {
			// Instance method: bind to constructor.prototype
			c.emitJSMemberGet(c.addConstant(NewString("prototype")))
		}

		// Compile the method/accessor body
		c.compileJScriptFunctionLiteral(md.Body, "", false)

		// Get property name
		var name string
		if id, ok := md.Key.(*jsast.Identifier); ok {
			name = id.Name.String()
		} else if lit, ok := md.Key.(*jsast.StringLiteral); ok {
			name = lit.Value.String()
		} else if lit, ok := md.Key.(*jsast.NumberLiteral); ok {
			name = fmt.Sprintf("%v", lit.Value)
		}

		nameIdx := c.addConstant(NewString(name))
		kind := jsPropertyKindMethod
		switch md.Kind {
		case jsast.PropertyKindGet:
			kind = jsPropertyKindGet
		case jsast.PropertyKindSet:
			kind = jsPropertyKindSet
		}

		c.emit(OpJSDefineProperty, nameIdx, kind)
	}

	// 3. Static Fields
	for _, el := range node.Body {
		if field, ok := el.(*jsast.FieldDefinition); ok && field.Static {
			c.emit(OpJSDup) // constructor
			if field.Initializer != nil {
				c.compileJScriptExpression(field.Initializer)
			} else {
				c.emit(OpJSLoadUndefined)
			}
			name := ""
			if id, ok := field.Key.(*jsast.Identifier); ok {
				name = id.Name.String()
			} else if id, ok := field.Key.(*jsast.PrivateIdentifier); ok {
				name = "\x00__priv_" + id.Name.String()
			} else if lit, ok := field.Key.(*jsast.StringLiteral); ok {
				name = lit.Value.String()
			}
			if name != "" {
				c.emitJSMemberSet(c.addConstant(NewString(name)))
			}
		}
	}
}

func (c *Compiler) compileJScriptClassFields() {
	for _, el := range c.jsClassFields {
		if field, ok := el.(*jsast.FieldDefinition); ok {
			c.emit(OpJSLoadThis)
			if field.Initializer != nil {
				c.compileJScriptExpression(field.Initializer)
			} else {
				c.emit(OpJSLoadUndefined)
			}
			name := ""
			if id, ok := field.Key.(*jsast.Identifier); ok {
				name = id.Name.String()
			} else if id, ok := field.Key.(*jsast.PrivateIdentifier); ok {
				name = "\x00__priv_" + id.Name.String()
			} else if lit, ok := field.Key.(*jsast.StringLiteral); ok {
				name = lit.Value.String()
			}
			if name != "" {
				c.emitJSMemberSet(c.addConstant(NewString(name)))
			}
		}
	}
}

// compileJScriptDefaultParamGuards emits bytecode at the beginning of a function
// body that checks each parameter with a default value and assigns the default
// expression when the actual argument was not provided (i.e., is undefined).
func (c *Compiler) compileJScriptDefaultParamGuards(paramList *jsast.ParameterList) {
	if paramList == nil {
		return
	}
	for _, b := range paramList.List {
		if b == nil || b.Initializer == nil {
			continue
		}
		p, ok := b.Target.(*jsast.Identifier)
		if !ok {
			continue
		}
		paramName := p.Name.String()
		nameIdx := c.addConstant(NewString(paramName))
		localSlot, hasLocal := c.jsResolveLocalSlot(paramName)

		// if (param === undefined) { param = defaultExpr; }
		if hasLocal {
			c.emit(OpJSGetLocal, localSlot)
		} else {
			c.emit(OpJSGetName, nameIdx)
		}
		c.emit(OpJSLoadUndefined)
		c.emit(OpJSStrictEq)
		// JumpIfFalse skips the default assignment when param is NOT undefined.
		skipJump := c.emitJSJump(OpJSJumpIfFalse)
		c.compileJScriptExpression(b.Initializer)
		if hasLocal {
			c.emit(OpJSSetLocal, localSlot)
		} else {
			c.emit(OpJSSetName, nameIdx)
		}
		c.patchJSJump(skipJump)
	}
}

// foldJSExpr recursively attempts to fold a JScript AST expression to a
// simpler literal at compile time. It mutates BinaryExpression children
// in-place and returns either the original node or a new literal node.
func foldJSExpr(expr jsast.Expression) jsast.Expression {
	bin, ok := expr.(*jsast.BinaryExpression)
	if !ok {
		return expr
	}
	// Fold children first (depth-first), enabling chained folding.
	bin.Left = foldJSExpr(bin.Left)
	bin.Right = foldJSExpr(bin.Right)
	// Try to evaluate this binary expression at compile time.
	if folded := foldJSBinaryLiterals(bin.Operator, bin.Left, bin.Right); folded != nil {
		return folded
	}
	return bin
}

// foldJSBinaryLiterals evaluates a binary operation over two literal AST nodes
// at compile time. Returns a new literal expression, or nil if not foldable.
func foldJSBinaryLiterals(op jstoken.Token, left, right jsast.Expression) jsast.Expression {
	lf, lok := jsExprAsFloat(left)
	rf, rok := jsExprAsFloat(right)
	ls, lsok := jsExprAsString(left)
	rs, rsok := jsExprAsString(right)

	switch op {
	case jstoken.PLUS:
		// + with any string operand coerces both sides to string.
		if lsok || rsok {
			if !lsok {
				if !lok {
					return nil
				}
				ls = jsFloatToString(lf)
			}
			if !rsok {
				if !rok {
					return nil
				}
				rs = jsFloatToString(rf)
			}
			return &jsast.StringLiteral{Value: jsunistring.String(ls + rs)}
		}
		if lok && rok {
			return jsNewNumLiteral(lf + rf)
		}
	case jstoken.MINUS:
		if lok && rok {
			return jsNewNumLiteral(lf - rf)
		}
	case jstoken.MULTIPLY:
		if lok && rok {
			return jsNewNumLiteral(lf * rf)
		}
	case jstoken.SLASH:
		if lok && rok && rf != 0 {
			return jsNewNumLiteral(lf / rf)
		}
	case jstoken.REMAINDER:
		if lok && rok && rf != 0 {
			return jsNewNumLiteral(math.Mod(lf, rf))
		}
	}
	return nil
}

// jsExprAsFloat extracts a float64 from a JS NumberLiteral node.
func jsExprAsFloat(expr jsast.Expression) (float64, bool) {
	num, ok := expr.(*jsast.NumberLiteral)
	if !ok {
		return 0, false
	}
	switch v := num.Value.(type) {
	case int64:
		return float64(v), true
	case int:
		return float64(v), true
	case float64:
		return v, true
	}
	return 0, false
}

// jsExprAsString extracts a plain Go string from a JS StringLiteral node.
func jsExprAsString(expr jsast.Expression) (string, bool) {
	sl, ok := expr.(*jsast.StringLiteral)
	if !ok {
		return "", false
	}
	return sl.Value.String(), true
}

// jsNewNumLiteral constructs a JS NumberLiteral carrying a float64 result.
// Idx and Literal are left at zero/empty since this node is never used for
// error reporting.
func jsNewNumLiteral(f float64) *jsast.NumberLiteral {
	return &jsast.NumberLiteral{Value: f}
}

// jsFloatToString converts a float64 to its JS string representation,
// omitting the decimal point for integer-valued floats.
func jsFloatToString(f float64) string {
	if f == math.Trunc(f) && !math.IsInf(f, 0) && !math.IsNaN(f) {
		return strconv.FormatInt(int64(f), 10)
	}
	return strconv.FormatFloat(f, 'f', -1, 64)
}
