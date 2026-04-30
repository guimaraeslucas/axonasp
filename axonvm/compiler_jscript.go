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
	"regexp"
	"strings"

	"g3pix.com.br/axonasp/jscript"
	jsast "g3pix.com.br/axonasp/jscript/ast"
	jsparser "g3pix.com.br/axonasp/jscript/parser"
	jstoken "g3pix.com.br/axonasp/jscript/token"
)

var jscriptCallAssignmentPattern = regexp.MustCompile(`([A-Za-z_][A-Za-z0-9_]*)\s*\(([^\)]*)\)\s*=\s*([^;\r\n]+);`)

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
		hasLexical := jsBlockHasLexicalDeclarations(node.List)
		if hasLexical {
			c.emit(OpJSBlockScopeEnter)
		}
		for i := range node.List {
			c.compileJScriptStatement(node.List[i])
		}
		if hasLexical {
			c.emit(OpJSBlockScopeExit)
		}
	case *jsast.TryStatement:
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
					c.emit(OpJSConstDeclare, nameIdx)
					if binding.Initializer != nil {
						c.compileJScriptExpression(binding.Initializer)
						c.emitConstInitialize(nameIdx)
					}
				} else {
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

	// Compile test condition (jump out if false)
	var jumpExit int
	if node.Test != nil {
		c.compileJScriptExpression(node.Test)
		jumpExit = c.emitJSJump(OpJSJumpIfFalse)
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
		c.compileJScriptExpression(node.Update)
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
			c.emit(OpJSConstDeclare, nameIdx)
		} else if isLexical {
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
		default:
			c.compileJScriptExpression(node.Left)
			c.compileJScriptExpression(node.Right)
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
					continue
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
		for i := range node.Value {
			if node.Value[i] == nil {
				c.emit(OpJSLoadUndefined)
				continue
			}
			c.compileJScriptExpression(node.Value[i])
		}
		c.emit(OpJSNewArray, len(node.Value))
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
	case *jsast.ConditionalExpression:
		c.compileJScriptExpression(node.Test)
		jumpFalse := c.emitJSJump(OpJSJumpIfFalse)
		c.compileJScriptExpression(node.Consequent)
		jumpEnd := c.emitJSJump(OpJSJump)
		c.patchJSJump(jumpFalse)
		c.compileJScriptExpression(node.Alternate)
		c.patchJSJump(jumpEnd)
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
		switch node.Operator {
		case jstoken.ASSIGN:
			c.emit(OpJSSetName, nameIdx)
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
		default:
			c.emit(OpJSSetName, nameIdx)
		}
		c.emit(OpJSLoadUndefined)
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

func (c *Compiler) compileJScriptCall(node *jsast.CallExpression) {
	switch callee := node.Callee.(type) {
	case *jsast.DotExpression:
		c.compileJScriptExpression(callee.Left)
		for i := range node.ArgumentList {
			c.compileJScriptExpression(node.ArgumentList[i])
		}
		c.emit(OpJSCallMember, c.addConstant(NewString(callee.Identifier.Name.String())), len(node.ArgumentList))
	default:
		c.compileJScriptExpression(node.Callee)
		for i := range node.ArgumentList {
			c.compileJScriptExpression(node.ArgumentList[i])
		}
		c.emit(OpJSCall, len(node.ArgumentList))
	}
}

func (c *Compiler) compileJScriptFunctionLiteral(fn *jsast.FunctionLiteral, fallbackName string) {
	jumpOverBody := c.emitJSJump(OpJSJump)
	bodyStart := len(c.bytecode)
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

// jsBlockHasLexicalDeclarations returns true if the statement list contains any
// LexicalDeclaration (let or const). Used to decide when to emit block scope opcodes.
func jsBlockHasLexicalDeclarations(stmts []jsast.Statement) bool {
	for _, s := range stmts {
		if _, ok := s.(*jsast.LexicalDeclaration); ok {
			return true
		}
	}
	return false
}

// compileJScriptLexicalDeclaration emits block-scoped let/const declarations.
func (c *Compiler) compileJScriptLexicalDeclaration(node *jsast.LexicalDeclaration) {
	isConst := node.Token == jstoken.CONST
	for _, binding := range node.List {
		name, ok := jsBindingIdentifierName(binding.Target)
		if !ok {
			continue
		}
		// Restricted identifiers (eval, arguments) are not allowed in strict mode
		if c.jsStrictMode && jsIsRestrictedIdentifier(name) {
			jsErr := jscript.NewJSSyntaxError(jscript.IllegalAssignment, 0, 0)
			jsErr.WithASPDescription(fmt.Sprintf("cannot use '%s' as a variable name in strict mode", name))
			if c.sourceName != "" {
				jsErr.WithFile(c.sourceName)
			}
			panic(jsErr)
		}
		nameIdx := c.addConstant(NewString(name))
		if isConst {
			c.emit(OpJSConstDeclare, nameIdx)
			if binding.Initializer != nil {
				c.compileJScriptExpression(binding.Initializer)
				c.emitConstInitialize(nameIdx)
			}
			// const without initializer is a SyntaxError per spec, but parser may allow it;
			// if initializer is nil, the const stays in TDZ indefinitely.
		} else {
			c.emit(OpJSLetDeclare, nameIdx)
			if binding.Initializer != nil {
				c.compileJScriptExpression(binding.Initializer)
				c.emit(OpJSSetName, nameIdx)
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
