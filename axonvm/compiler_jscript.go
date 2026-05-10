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

	program, err := jsparser.ParseFile(nil, c.sourceName, source, 0)
	if err != nil {
		panic(c.newJScriptCompileErrorFromParse(err, "jscript parse error"))
	}

	// Detect "use strict" directive at the beginning of the program
	hasStrictMode, _ := c.detectUseStrictDirective(program.Body)
	if hasStrictMode {
		c.emit(OpJSStrictModeEnter)
		prevStrictMode := c.jsStrictMode
		c.jsStrictMode = true
		defer func() { c.jsStrictMode = prevStrictMode }()
	}

	for i := range program.Body {
		c.compileJScriptStatement(program.Body[i])
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

	lastIdx := len(program.Body) - 1
	for i := 0; i < lastIdx; i++ {
		c.compileJScriptStatement(program.Body[i])
	}

	last := program.Body[lastIdx]
	if exprStmt, ok := last.(*jsast.ExpressionStatement); ok {
		c.compileJScriptExpression(exprStmt.Expression)
	} else {
		c.compileJScriptStatement(last)
		c.emit(OpJSLoadUndefined)
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

func (c *Compiler) compileJScriptStatement(stmt jsast.Statement) {
	switch node := stmt.(type) {
	case *jsast.ExpressionStatement:
		c.compileJScriptExpression(node.Expression)
		c.emit(OpJSPop)
	case *jsast.VariableStatement:
		for _, binding := range node.List {
			if binding.Initializer != nil {
				c.compileJScriptExpression(binding.Initializer)
				c.compileJScriptDestructuring(binding.Target, false, false, true)
			} else {
				// var x; -> declare x
				if id, ok := binding.Target.(*jsast.Identifier); ok {
					nameIdx := c.addConstant(NewString(id.Name.String()))
					c.emit(OpJSDeclareName, nameIdx)
				}
			}
		}
	case *jsast.LexicalDeclaration:
		// Handle ES6 let/const declarations with block scoping
		c.compileJScriptLexicalDeclaration(node)
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
		c.compileJScriptFunctionLiteral(node.Function, name)
		c.emit(OpJSSetName, nameIdx)
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
		if hasLexical {
			c.emit(OpJSBlockScopeEnter)
			for _, name := range letNames {
				c.emit(OpJSTDZRegisterLet, c.addConstant(NewString(name)))
			}
			for _, name := range constNames {
				c.emit(OpJSTDZRegisterConst, c.addConstant(NewString(name)))
			}
		}
		for i := range node.List {
			c.compileJScriptStatement(node.List[i])
		}
		if hasLexical {
			c.emit(OpJSBlockScopeExit)
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
				nameIdx := c.addConstant(NewString(id.Name.String()))
				c.emit(OpJSDeclareName, nameIdx)
				c.emit(OpJSLoadCatchError)
				c.emit(OpJSSetName, nameIdx)
			}
			c.compileJScriptStatement(node.Catch.Body)
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
		c.emitForIterExit(c.jsForIterScopes[i].nameIdxs)
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

	// Track whether we have a lexical (let/const) for-loop declaration
	var forIterNameIdxs []int
	isLexicalFor := false

	// Compile init expression
	if node.Initializer != nil {
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
				c.emit(OpJSDeclareName, nameIdx)
				if binding.Initializer != nil {
					c.compileJScriptExpression(binding.Initializer)
					c.emit(OpJSSetName, nameIdx)
				}
			}
		case *jsast.ForLoopInitializerLexicalDecl:
			// ES6 let/const for-loop: create outer block scope for the loop variable
			lexDecl := init.LexicalDeclaration
			isLexicalFor = lexDecl.Token == jstoken.LET
			c.emit(OpJSBlockScopeEnter)
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
		if fastNameIdx, fastLimitIdx, ok := c.detectJSForFastLessTest(node.Test); ok {
			jumpExit = c.emitJSJumpIfLessFast(fastNameIdx, fastLimitIdx)
		} else {
			c.compileJScriptExpression(node.Test)
			jumpExit = c.emitJSJump(OpJSJumpIfFalse)
		}
	}

	// For let loops: enter per-iteration scope by copying loop vars into a child env frame
	if isLexicalFor && len(forIterNameIdxs) > 0 {
		c.emitForIterEnter(forIterNameIdxs)
		c.jsForIterScopes = append(c.jsForIterScopes, jsForIterScope{nameIdxs: forIterNameIdxs})
	}

	// Compile loop body
	if node.Body != nil {
		c.compileJScriptStatement(node.Body)
	}

	// For let loops: exit per-iteration scope (write back updated vars to outer block scope).
	// This must be emitted BEFORE setting updateTarget so that continue statements (which
	// also emit ForIterExit via emitJSLeaveForIterScopes) jump to AFTER the ForIterExit
	// rather than triggering a double exit.
	if isLexicalFor && len(forIterNameIdxs) > 0 {
		c.jsForIterScopes = c.jsForIterScopes[:len(c.jsForIterScopes)-1]
		c.emitForIterExit(forIterNameIdxs)
	}

	// Mark update target: continue statements jump here (after per-iteration scope exit).
	updateTarget := len(c.bytecode)

	// Compile update expression
	if node.Update != nil {
		if !c.compileJScriptForUpdateFastPath(node.Update) {
			c.compileJScriptExpression(node.Update)
		}
		c.emit(OpJSPop)
	}

	// Jump back to test
	c.emitJSJumpTo(OpJSJump, loopCtx.loopStart)

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
	if node.Initializer != nil {
		if _, ok := node.Initializer.(*jsast.ForLoopInitializerLexicalDecl); ok {
			if len(c.jsBlockScopeStack) > 0 {
				c.jsBlockScopeStack = c.jsBlockScopeStack[:len(c.jsBlockScopeStack)-1]
			}
			c.emit(OpJSBlockScopeExit)
		}
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
		c.emit(OpJSGetName, c.addConstant(NewString(node.Name.String())))
	case *jsast.ThisExpression:
		c.emit(OpJSGetName, c.addConstant(NewString("this")))
	case *jsast.FunctionLiteral:
		c.compileJScriptFunctionLiteral(node, "")
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
			c.compileJScriptExpression(foldedLeft)
			c.compileJScriptExpression(foldedRight)
			switch node.Operator {
			case jstoken.PLUS:
				c.emit(OpJSAdd)
			case jstoken.MINUS:
				c.emit(OpJSSubtract)
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
	case *jsast.DotExpression:
		c.compileJScriptExpression(node.Left)
		c.emit(OpJSMemberGet, c.addConstant(NewString(node.Identifier.Name.String())))
	case *jsast.BracketExpression:
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
					c.emit(OpJSGetName, c.addConstant(NewString(key)))
				}
				c.emit(OpJSMemberSet, c.addConstant(NewString(key)))
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
					c.emit(OpJSMemberSet, c.addConstant(NewString(jsAccessorGetterPrefix+key)))
				case jsast.PropertyKindSet:
					c.emit(OpJSMemberSet, c.addConstant(NewString(jsAccessorSetterPrefix+key)))
				default:
					c.emit(OpJSMemberSet, c.addConstant(NewString(key)))
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
				c.emit(OpJSDelete, c.addConstant(NewString(t.Identifier.Name.String())))
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
		c.compileJScriptExpression(node.Right)
		if node.Operator == jstoken.ASSIGN {
			c.emit(OpJSDup)
			c.emit(OpJSSetName, nameIdx)
			return
		}
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
		case jstoken.EXPONENT_ASSIGN:
			c.emit(OpJSExponentAssign, nameIdx)
			return
		case jstoken.LOGICAL_AND_ASSIGN:
			c.emit(OpJSLogicalAndAssign, nameIdx)
			return
		case jstoken.LOGICAL_OR_ASSIGN:
			c.emit(OpJSLogicalOrAssign, nameIdx)
			return
		case jstoken.COALESCE_ASSIGN:
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
	case *jsast.DotExpression:
		c.compileJScriptExpression(left.Left)
		c.compileJScriptExpression(node.Right)
		c.emit(OpJSMemberSet, c.addConstant(NewString(left.Identifier.Name.String())))
		c.emit(OpJSLoadUndefined)
	case *jsast.BracketExpression:
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
func (c *Compiler) compileJScriptForUpdateFastPath(expr jsast.Expression) bool {
	assign, ok := expr.(*jsast.AssignExpression)
	if !ok {
		return false
	}

	leftID, ok := assign.Left.(*jsast.Identifier)
	if !ok {
		return false
	}

	name := leftID.Name.String()
	nameIdx := c.addConstant(NewString(name))

	// i += 1 / i -= 1
	if assign.Operator == jstoken.ADD_ASSIGN || assign.Operator == jstoken.SUBTRACT_ASSIGN {
		if !jsIsNumericOneLiteral(assign.Right) {
			return false
		}
		if assign.Operator == jstoken.ADD_ASSIGN {
			c.emit(OpJSPreIncrement, nameIdx)
		} else {
			c.emit(OpJSPreDecrement, nameIdx)
		}
		return true
	}

	// i = i + 1 / i = i - 1
	if assign.Operator != jstoken.ASSIGN {
		return false
	}

	rightBin, ok := assign.Right.(*jsast.BinaryExpression)
	if !ok {
		return false
	}
	rightLeftID, ok := rightBin.Left.(*jsast.Identifier)
	if !ok || rightLeftID.Name.String() != name {
		return false
	}
	if !jsIsNumericOneLiteral(rightBin.Right) {
		return false
	}

	if rightBin.Operator == jstoken.PLUS {
		c.emit(OpJSPreIncrement, nameIdx)
		return true
	}
	if rightBin.Operator == jstoken.MINUS {
		c.emit(OpJSPreDecrement, nameIdx)
		return true
	}

	return false
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

func (c *Compiler) compileJScriptCall(node *jsast.CallExpression) {
	switch callee := node.Callee.(type) {
	case *jsast.DotExpression:
		c.compileJScriptExpression(callee.Left)
		for i := range node.ArgumentList {
			c.compileJScriptExpression(node.ArgumentList[i])
		}
		c.emit(OpJSCallMember, c.addConstant(NewString(callee.Identifier.Name.String())), len(node.ArgumentList))
	case *jsast.BracketExpression:
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

func (c *Compiler) compileJScriptFunctionLiteral(fn *jsast.FunctionLiteral, fallbackName string) {
	jumpOverBody := c.emitJSJump(OpJSJump)
	bodyStart := len(c.bytecode)

	// Emit default parameter guards before the function body.
	c.compileJScriptDefaultParamGuards(fn.ParameterList)

	if fn.Body != nil {
		for i := range fn.Body.List {
			c.compileJScriptStatement(fn.Body.List[i])
		}
	}
	c.emit(OpJSLoadUndefined)
	c.emit(OpJSReturn)
	bodyEnd := len(c.bytecode)
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
		nameIdx := c.addConstant(NewString(operand.Name.String()))
		switch node.Operator {
		case jstoken.INCREMENT:
			if node.Postfix {
				c.emit(OpJSPostIncrement, nameIdx)
			} else {
				c.emit(OpJSPreIncrement, nameIdx)
			}
			return true
		case jstoken.DECREMENT:
			if node.Postfix {
				c.emit(OpJSPostDecrement, nameIdx)
			} else {
				c.emit(OpJSPreDecrement, nameIdx)
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
		}
	}
	return letNames, constNames
}

func (c *Compiler) compileJScriptDestructuring(target jsast.BindingTarget, isConst bool, isLet bool, isVar bool) {
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
			c.emit(OpJSDeclareName, nameIdx)
			c.emit(OpJSSetName, nameIdx)
		} else if isConst {
			c.emitConstInitialize(nameIdx)
		} else if isLet {
			c.emit(OpJSLetDeclare, nameIdx)
			c.emit(OpJSSetName, nameIdx)
		} else {
			// Normal assignment
			c.emit(OpJSSetName, nameIdx)
		}
	case *jsast.ObjectPattern:
		c.emit(OpJSRequireObject)
		for _, prop := range t.Properties {
			switch p := prop.(type) {
			case *jsast.PropertyShort:
				name := p.Name.Name.String()
				nameIdx := c.addConstant(NewString(name))
				c.emit(OpJSDup)
				c.emit(OpJSMemberGet, nameIdx)
				// TODO Phase 5.4: Default value
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
					c.emit(OpJSMemberGet, c.addConstant(NewString(key)))
				}
				// TODO Phase 5.4: Default value
				if bt, ok := p.Value.(jsast.BindingTarget); ok {
					c.compileJScriptDestructuring(bt, isConst, isLet, isVar)
				} else {
					c.emit(OpJSPop)
				}
			}
		}
		if t.Rest != nil {
			// TODO Phase 5.4: Object Rest
		}
		c.emit(OpJSPop) // Pop the source object
	case *jsast.ArrayPattern:
		// Phase 5.3: Array Destructuring
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
				c.compileJScriptDestructuring(binding.Target, true, false, false)
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
				c.compileJScriptDestructuring(binding.Target, false, true, false)
			} else {
				// let x; -> declare x
				if id, ok := binding.Target.(*jsast.Identifier); ok {
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

		// if (param === undefined) { param = defaultExpr; }
		c.emit(OpJSGetName, nameIdx)
		c.emit(OpJSLoadUndefined)
		c.emit(OpJSStrictEq)
		// JumpIfFalse skips the default assignment when param is NOT undefined.
		skipJump := c.emitJSJump(OpJSJumpIfFalse)
		c.compileJScriptExpression(b.Initializer)
		c.emit(OpJSSetName, nameIdx)
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
