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

	jsast "g3pix.com.br/axonasp/jscript/ast"
	jsparser "g3pix.com.br/axonasp/jscript/parser"
	jstoken "g3pix.com.br/axonasp/jscript/token"
	"g3pix.com.br/axonasp/vbscript"
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
		detail := err.Error()
		line, col := 0, 0
		if parserErr, ok := err.(*jsparser.Error); ok {
			line = parserErr.Position.Line
			col = parserErr.Position.Column
			detail = parserErr.Message
		} else if parserList, ok := err.(jsparser.ErrorList); ok && len(parserList) > 0 {
			line = parserList[0].Position.Line
			col = parserList[0].Position.Column
			detail = parserList[0].Message
		}

		vbErr := vbscript.NewVBSyntaxError(vbscript.SyntaxError, line, col, "", "")
		vbErr.WithASPDescription("jscript parse error: " + detail)
		if c.sourceName != "" {
			vbErr.WithFile(c.sourceName)
		}
		panic(vbErr)
	}
	for i := range program.Body {
		c.compileJScriptStatement(program.Body[i])
	}
}

// compileJScriptEvalSnippet parses one JScript eval source and emits OpJS bytecode
// that leaves the completion value on the stack and terminates with OpHalt.
func (c *Compiler) compileJScriptEvalSnippet(source string) {
	source = jscriptCallAssignmentPattern.ReplaceAllString(source, `$1($2, $3);`)

	program, err := jsparser.ParseFile(nil, c.sourceName, source, 0)
	if err != nil {
		panic(c.vbCompileError(vbscript.SyntaxError, fmt.Sprintf("jscript eval parse error: %v", err)))
	}

	if len(program.Body) == 0 {
		c.emit(OpJSLoadUndefined)
		c.emit(OpHalt)
		return
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

	c.emit(OpHalt)
}

// pushJSLoopContext adds a new loop context to the stack.
func (c *Compiler) pushJSLoopContext() *jsLoopContext {
	ctx := &jsLoopContext{
		continueTargets: make([]int, 0),
		loopStart:       len(c.bytecode),
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
	ctx := &jsBreakContext{breakTargets: make([]int, 0)}
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
		if node.Argument != nil {
			c.compileJScriptExpression(node.Argument)
		} else {
			c.emit(OpJSLoadUndefined)
		}
		c.emit(OpJSReturn)
	case *jsast.ThrowStatement:
		if node.Argument != nil {
			c.compileJScriptExpression(node.Argument)
		} else {
			c.emit(OpJSLoadUndefined)
		}
		c.emit(OpJSThrow)
	case *jsast.BlockStatement:
		for i := range node.List {
			c.compileJScriptStatement(node.List[i])
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
	case *jsast.BranchStatement:
		// BranchStatement handles both break and continue
		switch node.Token {
		case jstoken.BREAK:
			breakCtx := c.currentJSBreakContext()
			if breakCtx != nil {
				jumpPos := c.emitJSJump(OpJSBreak)
				breakCtx.breakTargets = append(breakCtx.breakTargets, jumpPos)
			}
		case jstoken.CONTINUE:
			loopCtx := c.currentJSLoopContext()
			if loopCtx != nil {
				jumpPos := c.emitJSJump(OpJSContinue)
				loopCtx.continueTargets = append(loopCtx.continueTargets, jumpPos)
			}
		}
	case *jsast.SwitchStatement:
		c.compileJScriptSwitchStatement(node)
	}
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
			// TODO: Handle lexical declaration (const/let)
		}
	}

	loopCtx.loopStart = len(c.bytecode)

	// Compile test condition (jump out if false)
	var jumpExit int
	if node.Test != nil {
		c.compileJScriptExpression(node.Test)
		jumpExit = c.emitJSJump(OpJSJumpIfFalse)
	}

	// Compile loop body
	if node.Body != nil {
		c.compileJScriptStatement(node.Body)
	}

	// Mark update target (for continue)
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
	// Patch continue targets to update
	for _, contPos := range loopCtx.continueTargets {
		c.patchJSJumpTo(contPos, updateTarget)
	}

	c.popJSLoopContext()
	c.popJSBreakContext()
}

// compileJScriptForInStatement compiles: for (var in object) statement
func (c *Compiler) compileJScriptForInStatement(node *jsast.ForInStatement) {
	loopCtx := c.pushJSLoopContext()
	breakCtx := c.pushJSBreakContext()

	varName := ""
	declareName := false
	switch into := node.Into.(type) {
	case *jsast.ForIntoVar:
		if into.Binding != nil {
			if name, ok := jsBindingIdentifierName(into.Binding.Target); ok {
				varName = name
				declareName = true
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
	if declareName {
		c.emit(OpJSDeclareName, nameIdx)
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
		nameIdx := c.addConstant(NewString(left.Name.String()))
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
		for _, b := range fn.ParameterList.List {
			if b == nil || b.Target == nil {
				continue
			}
			if p, ok := b.Target.(*jsast.Identifier); ok {
				params = append(params, p.Name.String())
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
