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
	"strings"
	"time"

	"g3pix.com.br/axonasp/vbscript"
)

type classPropertyAccessorKind int

const (
	classPropertyAccessorGet classPropertyAccessorKind = iota
	classPropertyAccessorLet
	classPropertyAccessorSet
)

// pushLoopContext registers one loop kind for Exit <kind> patching in nested scopes.
func (c *Compiler) pushLoopContext(kind string) {
	c.loopContexts = append(c.loopContexts, loopContext{kind: kind, exitJumps: make([]int, 0, 2)})
}

// appendLoopExitJump records one pending jump for Exit <kind> against the nearest enclosing loop kind.
func (c *Compiler) appendLoopExitJump(kind string, jumpPos int) {
	for i := len(c.loopContexts) - 1; i >= 0; i-- {
		if c.loopContexts[i].kind == kind {
			c.loopContexts[i].exitJumps = append(c.loopContexts[i].exitJumps, jumpPos)
			return
		}
	}
	if kind == "for" {
		panic(c.vbCompileError(vbscript.SyntaxError, "Exit For used outside For...Next"))
	}
	panic(c.vbCompileError(vbscript.SyntaxError, "Exit Do used outside Do...Loop"))
}

// popLoopContextAndPatch patches all recorded Exit jumps for the innermost loop context.
func (c *Compiler) popLoopContextAndPatch(loopEnd int) {
	if len(c.loopContexts) == 0 {
		return
	}
	ctx := c.loopContexts[len(c.loopContexts)-1]
	c.loopContexts = c.loopContexts[:len(c.loopContexts)-1]
	for _, jumpPos := range ctx.exitJumps {
		c.patchJumpTo(jumpPos, loopEnd)
	}
}

func (c *Compiler) parseStatement() {
	if c.matchEof() {
		return
	}

	c.emitCurrentDebugLocation()

	switch t := c.next.(type) {
	case *vbscript.HTMLToken:
		idx := c.addConstant(NewString(t.Content))
		c.emit(OpWriteStatic, idx)
		c.move()
		return
	case *vbscript.ASPExpressionStartToken:
		c.move()
		c.emitCurrentDebugLocation()
		c.parseExpression(PrecNone)
		c.emit(OpWrite)
		return
	case *vbscript.ASPDirectiveStartToken:
		c.move()
		c.compileASPDirective()
		return
	case *vbscript.ASPIncludeToken, *vbscript.ASPObjectToken:
		if objectToken, ok := t.(*vbscript.ASPObjectToken); ok {
			c.move()
			c.compileASPObjectDeclaration(objectToken)
			return
		}
		c.move()
		return
	case *vbscript.ASPCodeStartToken, *vbscript.ASPCodeEndToken:
		c.move()
		return
	case *vbscript.LineTerminationToken, *vbscript.ColonLineTerminationToken, *vbscript.CommentToken:
		c.move()
		return
	case *vbscript.KeywordToken:
		switch t.Keyword {
		case vbscript.KeywordOption:
			c.move() // Consume 'Option'
			c.parseOptionStatementAfterOptionKeyword()
			return
		case vbscript.KeywordOn:
			c.move()
			if c.matchKeywordOrIdentifier(vbscript.KeywordError, "error") {
				c.move()
				if c.matchKeywordOrIdentifier(vbscript.KeywordResume, "resume") {
					c.move()
					if c.checkKeyword(vbscript.KeywordNext) || c.matchKeywordOrIdentifier(vbscript.KeywordNext, "next") {
						c.move()
						c.emit(OpOnErrorResumeNext)
						return
					}
					panic(c.vbCompileError(vbscript.ExpectedNext, "Expected 'Next' after 'On Error Resume'"))
				} else if c.matchKeywordOrIdentifier(vbscript.KeywordGoto, "goto") {
					c.move()
					if lit, ok := c.next.(*vbscript.DecIntegerLiteralToken); ok && lit.Value == 0 {
						c.move()
						c.emit(OpOnErrorGoto0)
						return
					}
					panic(c.vbCompileError(vbscript.SyntaxError, "Expected '0' after 'On Error GoTo'"))
				}
				panic(c.vbCompileError(vbscript.SyntaxError, "Expected 'Resume Next' or 'GoTo 0' after 'On Error'"))
			}
			panic(c.vbCompileError(vbscript.SyntaxError, "Expected 'Error' after 'On'"))
		case vbscript.KeywordDim:
			c.parseDimStatement()
		case vbscript.KeywordEnum:
			c.parseEnumStatement()
		case vbscript.KeywordStatic:
			if !c.isLocal {
				panic(c.vbCompileError(vbscript.InvalidProcedureCallOrArgument, "'Static' is only valid within procedures"))
			}
			c.parseStaticStatement()
		case vbscript.KeywordConst:
			c.parseConstStatement()
		case vbscript.KeywordErase:
			c.parseEraseStatement()
		case vbscript.KeywordReDim:
			c.parseReDimStatement()
		case vbscript.KeywordIf:
			c.parseIfStatement()
		case vbscript.KeywordSelect:
			c.parseSelectCaseStatement()
		case vbscript.KeywordWhile:
			c.parseWhileStatement()
		case vbscript.KeywordDo:
			c.parseDoStatement()
		case vbscript.KeywordFor:
			c.parseForStatement()
		case vbscript.KeywordRaiseEvent:
			c.parseRaiseEventStatement()
		case vbscript.KeywordGoto:
			c.move()
			var name string
			if lit, ok := c.next.(*vbscript.DecIntegerLiteralToken); ok {
				name = fmt.Sprintf("%d", lit.Value)
				c.move()
			} else {
				name = c.expectIdentifier()
			}
			c.emitGoTo(name)
			return
		case vbscript.KeywordSub:
			c.parseSubFunction(false)
		case vbscript.KeywordFunction:
			c.parseSubFunction(true)
		case vbscript.KeywordClass:
			c.parseClassDeclaration()
		case vbscript.KeywordPublic, vbscript.KeywordPrivate:
			c.move()
			if c.matchKeywordOrIdentifier(vbscript.KeywordClass, "class") {
				c.parseClassDeclaration()
				return
			}
			if c.matchKeywordOrIdentifier(vbscript.KeywordEnum, "enum") {
				c.parseEnumStatement()
				return
			}
			if c.matchKeywordOrIdentifier(vbscript.KeywordEnd, "type") || strings.EqualFold(c.nextIdentifierName(), "type") {
				c.parseTypeDeclaration()
				return
			}
			if c.matchKeywordOrIdentifier(vbscript.KeywordConst, "const") {
				c.parseConstStatement()
				return
			}
			if c.matchKeywordOrIdentifier(vbscript.KeywordSub, "sub") {
				c.parseSubFunction(false)
				return
			}
			if c.matchKeywordOrIdentifier(vbscript.KeywordFunction, "function") {
				c.parseSubFunction(true)
				return
			}
			if !c.isLocal && c.currentClassName == "" {
				if c.checkKeyword(vbscript.KeywordDim) {
					c.parseDimStatement()
					return
				}
				if c.isIdentifierLikeToken(c.next) {
					c.parseScopedVariableDeclaration()
					return
				}
			}
			panic(c.vbCompileError(vbscript.ExpectedSub, "Expected Sub or Function after scope modifier"))
		case vbscript.KeywordSet:
			c.move()
			name := c.expectIdentifier()

			if p, ok := c.next.(*vbscript.PunctuationToken); ok && p.Type == vbscript.PunctLParen {
				op, idx := c.resolveVar(name)
				c.emit(op, idx)
				argCount := c.parseParenArgumentList()
				if p, ok := c.next.(*vbscript.PunctuationToken); ok && p.Type == vbscript.PunctEqual {
					c.move()
					c.parseExpression(PrecNone)
					midx := c.addConstant(NewString(""))
					c.emit(OpArraySet, midx, argCount)
					return
				}
				if p, ok := c.next.(*vbscript.PunctuationToken); ok && p.Type == vbscript.PunctDot {
					c.emit(OpCall, argCount)
					c.parseSetMemberAssignmentChain()
					return
				}
				panic(c.vbCompileError(vbscript.ExpectedEqual, "Expected '=' in Set assignment"))
			}

			if p, ok := c.next.(*vbscript.PunctuationToken); ok && p.Type == vbscript.PunctDot {
				op, idx := c.resolveVar(name)
				c.emit(op, idx)
				c.parseSetMemberAssignmentChain()
				return
			}

			if c.isLocal && c.currentClassName != "" {
				if _, exists := c.locals.Get(name); !exists {
					if property, ok := c.getClassPropertyDeclaration(c.currentClassName, name); ok && property != nil && property.HasSet {
						if p, ok := c.next.(*vbscript.PunctuationToken); !ok || p.Type != vbscript.PunctEqual {
							panic(c.vbCompileError(vbscript.ExpectedEqual, "Expected '=' in Set assignment"))
						}
						c.move()
						c.emit(OpMe)
						c.parseExpression(PrecNone)
						c.undoTrailingCoerce()
						midx := c.addConstant(NewString(name))
						c.emit(OpMemberSetSet, midx)
						return
					}
				}
			}

			if p, ok := c.next.(*vbscript.PunctuationToken); !ok || p.Type != vbscript.PunctEqual {
				panic(c.vbCompileError(vbscript.ExpectedEqual, "Expected '=' in Set assignment"))
			}
			c.move()
			rhsStart := len(c.bytecode)
			rhsBareName := ""
			switch t := c.next.(type) {
			case *vbscript.IdentifierToken:
				rhsBareName = strings.TrimSpace(t.Name)
			case *vbscript.ExtendedIdentifierToken:
				rhsBareName = strings.TrimSpace(strings.TrimSuffix(strings.TrimPrefix(t.Name, "["), "]"))
			case *vbscript.KeywordOrIdentifierToken:
				rhsBareName = strings.TrimSpace(t.Name)
			}
			c.parseExpression(PrecNone)
			// "Set" assigns a raw object reference; strip any OpCoerceToValue
			// the expression compiler may have emitted for the RHS identifier.
			c.undoTrailingCoerce()
			if rhsBareName != "" && len(c.bytecode) == rhsStart+3 && c.globalZeroArgFuncs[strings.ToLower(rhsBareName)] {
				if OpCode(c.bytecode[rhsStart]) == OpGetGlobal || OpCode(c.bytecode[rhsStart]) == OpConstant {
					c.emit(OpCall, 0)
				}
			}
			if c.isLocal && c.currentClassName != "" && rhsBareName != "" {
				// Forward class methods/properties used as bare RHS in Set (e.g. "Set x = dict")
				// can be miscompiled as OpGetGlobal due single-pass ordering. When the RHS is
				// exactly one identifier-load opcode and the target global was never Dim/Const
				// declared by user code, remap it to class-member resolution.
				if len(c.bytecode) == rhsStart+3 && OpCode(c.bytecode[rhsStart]) == OpGetGlobal {
					lower := strings.ToLower(rhsBareName)
					if !c.declaredGlobals[lower] && !c.constGlobals[lower] {
						memberIdx := c.addConstant(NewString(rhsBareName))
						c.bytecode[rhsStart] = byte(OpGetClassMember)
						binary.BigEndian.PutUint16(c.bytecode[rhsStart+1:rhsStart+3], uint16(memberIdx))
					}
				}
			}
			op, idx := c.resolveSetVar(name)
			c.emit(op, idx)
		case vbscript.KeywordExit:
			c.move()
			if c.matchKeywordOrIdentifier(vbscript.KeywordSub, "sub") {
				// Exit Sub never has a return value.
				c.move()
				emptyIdx := c.addConstant(NewEmpty())
				c.emit(OpConstant, emptyIdx)
				c.emit(OpRet)
				return
			}
			if c.matchKeywordOrIdentifier(vbscript.KeywordFunction, "function") || c.matchKeywordOrIdentifier(vbscript.KeywordProperty, "property") {
				// Exit Function / Exit Property must return the current function return value,
				// not Empty. The return variable is a local slot named after the function.
				c.move()
				if c.currentFunctionName != "" {
					op, idx := c.resolveVar(c.currentFunctionName)
					c.emit(op, idx)
				} else {
					emptyIdx := c.addConstant(NewEmpty())
					c.emit(OpConstant, emptyIdx)
				}
				c.emit(OpRet)
				return
			}
			if c.matchKeywordOrIdentifier(vbscript.KeywordFor, "for") {
				c.move()
				exitJump := c.emitJump(OpJump)
				c.appendLoopExitJump("for", exitJump)
				return
			}
			if c.matchKeywordOrIdentifier(vbscript.KeywordDo, "do") {
				c.move()
				exitJump := c.emitJump(OpJump)
				c.appendLoopExitJump("do", exitJump)
				return
			}
			panic(c.vbCompileError(vbscript.ExpectedSub, "Expected Sub, Function, Property, For, or Do after Exit"))
		case vbscript.KeywordWith:
			c.parseWithStatement()
		case vbscript.KeywordCall:
			c.move()
			// Parse the callee (function/subroutine name or member access)
			c.parseExpression(PrecNone)
			c.emit(OpPop)
		case vbscript.KeywordOpen:
			c.parseOpenStatement()
		case vbscript.KeywordPrint:
			c.parsePrintStatement()
		case vbscript.KeywordWrite:
			c.parseWriteStatement()
		case vbscript.KeywordLine:
			c.move()
			if c.matchKeywordOrIdentifier(vbscript.KeywordInput, "input") {
				c.move()
				c.parseLineInputStatement()
			} else {
				panic(c.vbCompileError(vbscript.ExpectedIdentifier, "Expected 'Input' after 'Line'"))
			}
		case vbscript.KeywordInput:
			c.parseInputStatement()
		case vbscript.KeywordClose:
			c.parseCloseStatement()
		case vbscript.KeywordPut:
			c.parsePutStatement()
		case vbscript.KeywordGet:
			c.parseGetStatement()
		default:
			c.parseExpression(PrecNone)
			c.emit(OpPop)
		}
	case *vbscript.DecIntegerLiteralToken:
		// VB6 supports line numbers as labels
		lit := c.next.(*vbscript.DecIntegerLiteralToken)
		name := fmt.Sprintf("%d", lit.Value)
		c.move()

		if _, ok := c.next.(*vbscript.ColonLineTerminationToken); ok {
			c.move() // Consume ':'
			c.registerLabel(name)
		}
		return
	case *vbscript.IdentifierToken, *vbscript.ExtendedIdentifierToken, *vbscript.KeywordOrIdentifierToken:
		var name string
		switch t := c.next.(type) {
		case *vbscript.IdentifierToken:
			name = t.Name
		case *vbscript.ExtendedIdentifierToken:
			name = strings.TrimSuffix(strings.TrimPrefix(t.Name, "["), "]")
		case *vbscript.KeywordOrIdentifierToken:
			name = t.Name
		}

		c.move()

		if _, ok := c.next.(*vbscript.ColonLineTerminationToken); ok {
			c.move() // Consume ':'
			c.registerLabel(name)
			return
		}

		if strings.EqualFold(name, "Option") {
			c.parseOptionStatementAfterOptionKeyword()
			return
		}

		if strings.EqualFold(name, "Type") {
			c.parseTypeDeclaration()
			return
		}

		if strings.EqualFold(name, "Erase") {
			c.parseEraseStatementAfterNameToken()
			return
		}

		if strings.EqualFold(name, "Print") && !c.matchPunctuation(vbscript.PunctEqual) && !c.matchPunctuation(vbscript.PunctDot) && !c.matchPunctuation(vbscript.PunctLParen) {
			c.parsePrintStatementAfterPrint()
			return
		}
		if strings.EqualFold(name, "Write") && !c.matchPunctuation(vbscript.PunctEqual) && !c.matchPunctuation(vbscript.PunctDot) && !c.matchPunctuation(vbscript.PunctLParen) {
			c.parseWriteStatementAfterWrite()
			return
		}
		if strings.EqualFold(name, "Close") && !c.matchPunctuation(vbscript.PunctEqual) && !c.matchPunctuation(vbscript.PunctDot) && !c.matchPunctuation(vbscript.PunctLParen) {
			c.parseCloseStatementAfterClose()
			return
		}
		if strings.EqualFold(name, "Input") && !c.matchPunctuation(vbscript.PunctEqual) && !c.matchPunctuation(vbscript.PunctDot) && !c.matchPunctuation(vbscript.PunctLParen) {
			c.parseInputStatementAfterInput()
			return
		}
		if strings.EqualFold(name, "Line") && c.matchKeywordOrIdentifier(vbscript.KeywordInput, "input") {
			c.move()
			c.parseLineInputStatement()
			return
		}
		if strings.EqualFold(name, "Put") && !c.matchPunctuation(vbscript.PunctEqual) && !c.matchPunctuation(vbscript.PunctDot) && !c.matchPunctuation(vbscript.PunctLParen) {
			c.parsePutStatementAfterPut()
			return
		}
		if strings.EqualFold(name, "Get") && !c.matchPunctuation(vbscript.PunctEqual) && !c.matchPunctuation(vbscript.PunctDot) && !c.matchPunctuation(vbscript.PunctLParen) {
			c.parseGetStatementAfterGet()
			return
		}

		if p, ok := c.next.(*vbscript.PunctuationToken); ok && p.Type == vbscript.PunctEqual {
			c.move() // Consume '='
			op, idx := c.resolveSetVar(name)
			rhsStart := len(c.bytecode)
			c.parseExpression(PrecNone)
			if c.tryOptimizeGlobalIncrementAssignment(rhsStart, op, idx) {
				return
			}
			if c.isLocal && c.currentFunctionName != "" && strings.EqualFold(name, c.currentFunctionName) {
				// Function/Property return slot assignment must always overwrite the
				// local Variant directly, regardless of its previous runtime subtype.
				// Using Let-dispatch here can incorrectly try default-property writes.
				c.emit(op, idx)
				return
			}
			// Plain "name = value" uses Let opcodes to preserve VBScript's
			// non-Set assignment semantics while keeping variable-slot overwrites
			// distinct from explicit object-reference Set assignments.
			c.emit(c.letOpCode(op), idx)
		} else if p, ok := c.next.(*vbscript.PunctuationToken); ok && p.Type == vbscript.PunctLParen {
			if c.compileImplicitClassStatementCall(name, true) {
				return
			}
			// Detect forward function/sub calls inside local scopes.
			// Bare calls like "foo()" where foo is not yet a known local or
			// declared global must resolve as OpGetGlobal so the forward-call
			// patching mechanism (patchForwardCallSites) can rewrite the load
			// when the target Sub/Function is defined later in source.
			forceGlobal := false
			if c.isLocal && c.currentClassName == "" {
				lower := strings.ToLower(strings.TrimSpace(name))
				if _, localExists := c.locals.Get(name); !localExists {
					if _, isStatic := c.staticLocals[lower]; !isStatic {
						if _, globalExists := c.Globals.Get(name); !globalExists || !c.declaredGlobals[lower] {
							forceGlobal = true
						}
					}
				}
			}
			var op OpCode
			var idx int
			var loadPos int
			if forceGlobal {
				idx, exists := c.Globals.Get(name)
				if !exists {
					idx = c.Globals.Add(name)
				}
				op = OpGetGlobal
				loadPos = c.emit(op, idx)
			} else {
				op, idx = c.resolveVar(name)
				loadPos = c.emit(op, idx)
			}
			argCount := c.parseParenArgumentList()

			if peq, ok := c.next.(*vbscript.PunctuationToken); ok && peq.Type == vbscript.PunctEqual {
				c.move() // Consume '='
				c.parseExpression(PrecNone)
				midx := c.addConstant(NewString(""))
				c.emit(OpArraySet, midx, argCount)
			} else {
				c.emit(OpCall, argCount)
				if op == OpGetGlobal {
					c.registerForwardCallPatch(name, loadPos)
				}
				if c.parseStatementCallChain() {
					return
				}
				c.emit(OpPop)
			}
		} else if p, ok := c.next.(*vbscript.PunctuationToken); ok && p.Type == vbscript.PunctDot {
			// Member call or property set: Response.Write "Hello" / obj.Prop = value
			// Save the bytecode position BEFORE emitting the object-load opcode so that
			// the Response.Write peephole optimisation can trim it if the fast path applies.
			objectEmitStart := len(c.bytecode)
			if !c.emitStaticObjectIdentifierFallback(name) {
				op, idx := c.resolveVar(name)
				c.emit(op, idx)
			}

			c.move() // Consume "."
			memberName := c.expectIdentifier()
			memberChain := []string{memberName}
			for {
				if dot, ok := c.next.(*vbscript.PunctuationToken); ok && dot.Type == vbscript.PunctDot {
					c.move()
					memberName = c.expectIdentifier()
					memberChain = append(memberChain, memberName)
					continue
				}
				break
			}

			flatMemberName := strings.Join(memberChain, ".")
			callMemberName := flatMemberName
			if len(memberChain) > 1 {
				isAssignment := false
				if peq, ok := c.next.(*vbscript.PunctuationToken); ok && peq.Type == vbscript.PunctEqual {
					isAssignment = true
				}
				if !isAssignment {
					for i := 0; i < len(memberChain)-1; i++ {
						intermediateIdx := c.addConstant(NewString(memberChain[i]))
						c.emit(OpConstant, intermediateIdx)
						c.emit(OpMemberGet)
					}
					callMemberName = memberChain[len(memberChain)-1]
				}
			}

			// Property assignment: obj.Prop = value
			if peq, ok := c.next.(*vbscript.PunctuationToken); ok && peq.Type == vbscript.PunctEqual {
				// Check if target is a known UDT to use fast ExtOpSetRecordMember
				udtName, isUDT := c.lastEmittedUDTNameFromOp()
				if isUDT && len(memberChain) == 1 {
					memberIdx, _, _, found := c.getUDTMemberIndex(udtName, memberName)
					if found {
						c.move() // Consume '='
						c.parseExpression(PrecNone)
						c.emitExt(ExtOpSetRecordMember, memberIdx)
						return
					}
				}

				setMemberName := flatMemberName
				if len(memberChain) > 1 {
					for i := 0; i < len(memberChain)-1; i++ {
						intermediateIdx := c.addConstant(NewString(memberChain[i]))
						c.emit(OpConstant, intermediateIdx)
						c.emit(OpMemberGet)
					}
					setMemberName = memberChain[len(memberChain)-1]
				}
				c.move() // Consume '='
				c.parseExpression(PrecNone)
				midx := c.addConstant(NewString(setMemberName))
				c.emit(OpMemberSet, midx)
			} else if lp, ok := c.next.(*vbscript.PunctuationToken); ok && lp.Type == vbscript.PunctLParen {
				argCount := c.parseParenArgumentList()
				midx := c.addConstant(NewString(callMemberName))
				if peq, ok2 := c.next.(*vbscript.PunctuationToken); ok2 && peq.Type == vbscript.PunctEqual {
					c.move() // Consume '='
					c.parseExpression(PrecNone)
					c.emit(OpArraySet, midx, argCount)
				} else if dot, ok2 := c.next.(*vbscript.PunctuationToken); ok2 && dot.Type == vbscript.PunctDot {
					// Chained member access after call: obj.Method(args).Prop = value
					// or obj.Method(args).SubMethod(args). Calls parseStatementCallChain
					// which handles .Prop = value (OpMemberSet), .SubMethod() (OpCallMember),
					// .Prop (OpMemberGet) and terminates with OpPop if needed.
					c.emit(OpCallMember, midx, argCount)
					if c.parseStatementCallChain() {
						return
					}
					c.emit(OpPop)
				} else {
					c.emit(OpCallMember, midx, argCount)
					if c.parseStatementCallChain() {
						return
					}
					c.emit(OpPop)
				}
			} else {
				// Member sub call without parentheses: obj.Method arg1, arg2
				//
				// Peephole optimisation for Response.Write <expr>:
				// When the target is the intrinsic Response object (not shadowed by the user)

				// and the method is Write with a single argument that is a top-level &
				// concatenation chain, flatten the chain into individual stack pushes and
				// emit OpWriteN(N) instead of the normal OpConcat+OpCallMember sequence.
				// This eliminates all intermediate concatenated string Value allocations.
				// The optimisation is safe because all operands are fully evaluated before
				// any write occurs, preserving On Error Resume Next semantics.
				if !c.isStatementEnd() &&
					strings.EqualFold(name, "Response") &&
					len(memberChain) == 1 &&
					strings.EqualFold(memberChain[0], "Write") {
					// Trim the object-load opcode we already emitted for Response; OpWriteN
					// goes directly through vm.output so the object reference is not needed.
					c.bytecode = c.bytecode[:objectEmitStart]
					count := c.parseResponseWriteFlatChain()
					c.emit(OpWriteN, count)
					return
				}

				argCount := 0
				if !c.isStatementEnd() {
					for {
						if comma, ok := c.next.(*vbscript.PunctuationToken); ok && comma.Type == vbscript.PunctComma {
							emptyIdx := c.addConstant(NewEmpty())
							c.emit(OpConstant, emptyIdx)
						} else if c.isStatementEnd() {
							emptyIdx := c.addConstant(NewEmpty())
							c.emit(OpConstant, emptyIdx)
						} else {
							mscArgStartPos := len(c.bytecode)
							c.parseExpression(PrecNone)
							c.patchArgRefInBytecode(mscArgStartPos)
						}
						argCount++
						if p, ok := c.next.(*vbscript.PunctuationToken); ok && p.Type == vbscript.PunctComma {
							c.move()
						} else {
							break
						}
					}
				}
				midx := c.addConstant(NewString(callMemberName))
				c.emit(OpCallMember, midx, argCount)
				c.emit(OpPop)
			}
		} else {
			// Sub call without parens: MySub 1, 2
			if c.compileImplicitClassStatementCall(name, false) {
				return
			}
			// Determine BEFORE resolveVar whether the name is a known variable,
			// because resolveVar will add unknown names to implicitGlobals.
			lowerName := strings.ToLower(name)
			isKnownZeroArgSub := c.globalZeroArgSubs[lowerName] || c.globalZeroArgFuncs[lowerName]
			isKnownVariable := c.declaredGlobals[lowerName] || c.implicitGlobals[lowerName] || c.constGlobals[lowerName]
			if !isKnownVariable && c.isLocal {
				if _, lok := c.locals.Get(name); lok {
					isKnownVariable = true
				} else if _, sok := c.staticLocals[lowerName]; sok {
					isKnownVariable = true
				}
			}

			op, idx := c.resolveVar(name)
			loadPos := c.emit(op, idx)

			if !c.isStatementEnd() {
				argCount := 0
				for {
					if comma, ok := c.next.(*vbscript.PunctuationToken); ok && comma.Type == vbscript.PunctComma {
						emptyIdx := c.addConstant(NewEmpty())
						c.emit(OpConstant, emptyIdx)
					} else if c.isStatementEnd() {
						emptyIdx := c.addConstant(NewEmpty())
						c.emit(OpConstant, emptyIdx)
					} else {
						subNoParenStartPos := len(c.bytecode)
						c.parseExpression(PrecNone)
						c.patchArgRefInBytecode(subNoParenStartPos)
					}
					argCount++
					if p, ok := c.next.(*vbscript.PunctuationToken); ok && p.Type == vbscript.PunctComma {
						c.move()
					} else {
						break
					}
				}
				c.emit(OpCall, argCount)
				if op == OpGetGlobal {
					c.registerForwardCallPatch(name, loadPos)
				}
			} else if isKnownZeroArgSub || (!isKnownVariable && op == OpGetGlobal) {
				// Zero-argument bare Sub/Function call: the identifier is alone on the
				// statement line with no "=", "(", or "." following it. In VBScript this
				// invokes the Sub/Function, discarding any return value.
				//
				// Two cases:
				//   1. Backward reference — the Sub/Function was declared above the call
				//      site; its name is tracked in globalZeroArgSubs/globalZeroArgFuncs.
				//   2. Forward reference — the Sub/Function is declared below the call
				//      site and its name is not yet tracked. We emit OpCall(0)
				//      optimistically; patchForwardCallSites will rewrite the load when the
				//      declaration is parsed. If the target is never resolved as a
				//      Sub/Function, the runtime pushes VTEmpty for non-callable values.
				c.emit(OpCall, 0)
				if op == OpGetGlobal {
					c.registerForwardCallPatch(name, loadPos)
				}
			}
			c.emit(OpPop)
		}
	case *vbscript.PunctuationToken:
		// A statement beginning with '.' inside a With block: .Prop = val, .Method args
		if t.Type == vbscript.PunctDot && c.withDepth > 0 {
			c.move() // consume '.'
			c.compileWithMemberStatement()
			return
		}
		c.move()
	default:
		c.move()
	}
}

// parseStatementCallChain continues statement parsing after a call result is on the stack.
// It preserves chained default/member assignments like obj(i)(k) = v and only discards
// the final value when the chain is used as a plain statement call/read.
func (c *Compiler) parseStatementCallChain() bool {
	handled := false
	for {
		if dot, ok := c.next.(*vbscript.PunctuationToken); ok && dot.Type == vbscript.PunctDot {
			handled = true
			c.move()
			memberName := c.expectIdentifier()
			midx := c.addConstant(NewString(memberName))

			if lp, ok := c.next.(*vbscript.PunctuationToken); ok && lp.Type == vbscript.PunctLParen {
				argCount := c.parseParenArgumentList()
				if peq, ok := c.next.(*vbscript.PunctuationToken); ok && peq.Type == vbscript.PunctEqual {
					c.move()
					c.parseExpression(PrecNone)
					c.emit(OpArraySet, midx, argCount)
					return true
				}
				c.emit(OpCallMember, midx, argCount)
				continue
			}

			if peq, ok := c.next.(*vbscript.PunctuationToken); ok && peq.Type == vbscript.PunctEqual {
				c.move()
				c.parseExpression(PrecNone)
				c.emit(OpMemberSet, midx)
				return true
			}

			c.emit(OpConstant, midx)
			c.emit(OpMemberGet)
			continue
		}

		if lp, ok := c.next.(*vbscript.PunctuationToken); ok && lp.Type == vbscript.PunctLParen {
			handled = true
			argCount := c.parseParenArgumentList()
			if peq, ok := c.next.(*vbscript.PunctuationToken); ok && peq.Type == vbscript.PunctEqual {
				c.move()
				c.parseExpression(PrecNone)
				midx := c.addConstant(NewString(""))
				c.emit(OpArraySet, midx, argCount)
				return true
			}
			c.emit(OpCall, argCount)
			continue
		}

		break
	}

	if handled {
		c.emit(OpPop)
	}
	return handled
}

// tryOptimizeGlobalIncrementAssignment rewrites `name = name +/- 1` into the dedicated
// global increment/decrement opcode when the target resolves to a global slot.
func (c *Compiler) tryOptimizeGlobalIncrementAssignment(rhsStart int, op OpCode, idx int) bool {
	if op != OpSetGlobal || rhsStart < 0 || rhsStart >= len(c.bytecode) {
		return false
	}
	code := c.bytecode[rhsStart:]
	offset := 3
	if len(code) == 8 && OpCode(code[3]) == OpCoerceToValue {
		offset = 4
	} else if len(code) != 7 {
		return false
	}
	if OpCode(code[0]) != OpGetGlobal {
		return false
	}
	if int(binary.BigEndian.Uint16(code[1:3])) != idx {
		return false
	}
	if OpCode(code[offset]) != OpConstant {
		return false
	}
	constIdx := int(binary.BigEndian.Uint16(code[offset+1 : offset+3]))
	if constIdx < 0 || constIdx >= len(c.constants) {
		return false
	}
	constVal := c.constants[constIdx]
	if constVal.Type == VTInteger && constVal.Num == 1 {
		if OpCode(code[len(code)-1]) == OpAdd {
			c.bytecode = c.bytecode[:rhsStart]
			c.emit(OpIncGlobalInt, idx)
			return true
		}
		if OpCode(code[len(code)-1]) == OpSub {
			c.bytecode = c.bytecode[:rhsStart]
			c.emit(OpDecGlobalInt, idx)
			return true
		}
	}
	if constVal.Type == VTDouble && constVal.Flt == 1 {
		if OpCode(code[len(code)-1]) == OpAdd {
			c.bytecode = c.bytecode[:rhsStart]
			c.emit(OpIncGlobalInt, idx)
			return true
		}
		if OpCode(code[len(code)-1]) == OpSub {
			c.bytecode = c.bytecode[:rhsStart]
			c.emit(OpDecGlobalInt, idx)
			return true
		}
	}
	return false
}

// parseSetMemberAssignmentChain compiles Set assignments targeting chained member expressions.
// It walks each link explicitly (Obj.A.B, Obj.Collection(0).Value) and emits intermediate
// gets/calls instead of flattening names with literal dots.
func (c *Compiler) parseSetMemberAssignmentChain() {
	if p, ok := c.next.(*vbscript.PunctuationToken); !ok || p.Type != vbscript.PunctDot {
		panic(c.vbCompileError(vbscript.SyntaxError, "Expected '.' in Set member assignment"))
	}

	for {
		c.move() // consume '.'
		memberName := c.expectIdentifier()
		midx := c.addConstant(NewString(memberName))

		hasCall := false
		argCount := 0
		if lp, ok := c.next.(*vbscript.PunctuationToken); ok && lp.Type == vbscript.PunctLParen {
			hasCall = true
			argCount = c.parseParenArgumentList()
		}

		if dot, ok := c.next.(*vbscript.PunctuationToken); ok && dot.Type == vbscript.PunctDot {
			if hasCall {
				c.emit(OpCallMember, midx, argCount)
			} else {
				c.emit(OpConstant, midx)
				c.emit(OpMemberGet)
			}
			continue
		}

		if hasCall {
			panic(c.vbCompileError(vbscript.SyntaxError, "Expected member name after indexed call target in Set assignment"))
		}

		if p, ok := c.next.(*vbscript.PunctuationToken); !ok || p.Type != vbscript.PunctEqual {
			panic(c.vbCompileError(vbscript.ExpectedEqual, "Expected '=' in Set member assignment"))
		}

		c.move()
		c.parseExpression(PrecNone)
		// Set member assignment must preserve raw object references.
		c.undoTrailingCoerce()
		c.emit(OpMemberSetSet, midx)
		return
	}
}

// parseOptionStatementAfterOptionKeyword consumes Option sub-keywords after "Option" and updates compiler options.
func (c *Compiler) parseOptionStatementAfterOptionKeyword() {
	if c.checkKeyword(vbscript.KeywordExplicit) {
		c.move()
		c.optionExplicit = true
		return
	}

	if c.checkKeyword(vbscript.KeywordBase) {
		c.move()
		if lit, ok := c.next.(*vbscript.DecIntegerLiteralToken); ok {
			if lit.Value == 0 || lit.Value == 1 {
				c.optionBase = int(lit.Value)
				c.move()
				return
			}
		}
		panic(c.vbCompileError(vbscript.SyntaxError, "Expected '0' or '1' after 'Option Base'"))
	}

	if c.matchKeywordOrIdentifier(vbscript.KeywordCompare, "compare") {
		c.move()
		if c.matchKeywordOrIdentifier(vbscript.KeywordText, "text") {
			c.optionCompare = 1
			c.emit(OpSetOption, 0, 1) // Option Compare Text
			c.move()
			return
		}
		if c.matchKeywordOrIdentifier(vbscript.KeywordBinary, "binary") {
			c.optionCompare = 0
			c.emit(OpSetOption, 0, 0) // Option Compare Binary
			c.move()
			return
		}
		return
	}

	if id, ok := c.next.(*vbscript.IdentifierToken); ok {
		if strings.EqualFold(id.Name, "Infer") {
			c.move()
			if id2, ok := c.next.(*vbscript.IdentifierToken); ok && strings.EqualFold(id2.Name, "On") {
				c.optionInfer = true
				c.move()
			}
			return
		}
		if strings.EqualFold(id.Name, "Strict") {
			c.move()
			if id2, ok := c.next.(*vbscript.IdentifierToken); ok && strings.EqualFold(id2.Name, "On") {
				c.optionStrict = true
				c.move()
			}
		}
	}
}

// parseWithStatement compiles a With...End With block.
// The With subject expression is evaluated, stored on the VM with-object stack via OpWithEnter,
// the body statements are compiled (with c.withDepth > 0 so '.' syntax is enabled),
// and OpWithLeave restores the stack on exit.
func (c *Compiler) parseWithStatement() {
	c.expectKeyword(vbscript.KeywordWith)
	// Evaluate the With-subject expression and push it onto the data stack.
	c.parseExpression(PrecNone)
	// Move TOS object to the with-object stack; data stack is unchanged depth-wise.
	c.emit(OpWithEnter)
	c.withDepth++

	// Compile body until End With.
	for !c.matchEof() {
		if kw, ok := c.next.(*vbscript.KeywordToken); ok && kw.Keyword == vbscript.KeywordEnd {
			break
		}
		c.parseStatement()
	}

	c.expectKeyword(vbscript.KeywordEnd)
	c.expectKeyword(vbscript.KeywordWith)
	c.withDepth--
	c.emit(OpWithLeave)
}

// compileWithMemberStatement compiles a statement that starts with '.' inside a With block.
// The leading '.' token has already been consumed by the caller.
// Handles: .Prop = value, .Method(args), .Method args, and chained .A.B = value.
func (c *Compiler) compileWithMemberStatement() {
	// Push the innermost With-subject object onto the data stack.
	c.emit(OpWithLoad)

	memberName := c.expectIdentifier()

	// Collect chained access: .A.B.C — matches existing "name.A.B" member-set pattern.
	for {
		if dot, ok := c.next.(*vbscript.PunctuationToken); ok && dot.Type == vbscript.PunctDot {
			c.move()
			memberName = memberName + "." + c.expectIdentifier()
			continue
		}
		break
	}

	midx := c.addConstant(NewString(memberName))

	if peq, ok := c.next.(*vbscript.PunctuationToken); ok && peq.Type == vbscript.PunctEqual {
		// .Prop = value  –– implicit Let assignment
		c.move()
		c.parseExpression(PrecNone)
		c.emit(OpMemberSet, midx)
	} else if lp, ok := c.next.(*vbscript.PunctuationToken); ok && lp.Type == vbscript.PunctLParen {
		// .Method(args) or .Arr(idx) = value
		argCount := c.parseParenArgumentList()
		if peq, ok2 := c.next.(*vbscript.PunctuationToken); ok2 && peq.Type == vbscript.PunctEqual {
			c.move()
			c.parseExpression(PrecNone)
			c.emit(OpArraySet, midx, argCount)
		} else {
			c.emit(OpCallMember, midx, argCount)
			c.emit(OpPop)
		}
	} else {
		// .Method arg1, arg2  –– no-parens member call
		argCount := 0
		if !c.isStatementEnd() {
			for {
				if comma, ok := c.next.(*vbscript.PunctuationToken); ok && comma.Type == vbscript.PunctComma {
					emptyIdx := c.addConstant(NewEmpty())
					c.emit(OpConstant, emptyIdx)
				} else if c.isStatementEnd() {
					emptyIdx := c.addConstant(NewEmpty())
					c.emit(OpConstant, emptyIdx)
				} else {
					argStartPos := len(c.bytecode)
					c.parseExpression(PrecNone)
					c.patchArgRefInBytecode(argStartPos)
				}
				argCount++
				if p, ok := c.next.(*vbscript.PunctuationToken); ok && p.Type == vbscript.PunctComma {
					c.move()
				} else {
					break
				}
			}
		}
		c.emit(OpCallMember, midx, argCount)
		c.emit(OpPop)
	}
}

// registerClassDeclaration stores a class-name stub for staged class implementation.
func (c *Compiler) registerClassDeclaration(name string) {
	trimmedName := strings.TrimSpace(name)
	if trimmedName == "" {
		return
	}

	lowerName := strings.ToLower(trimmedName)
	if _, exists := c.classDeclLookup[lowerName]; exists {
		return
	}

	c.classDeclLookup[lowerName] = len(c.classDecls)
	c.classDecls = append(c.classDecls, CompiledClassDecl{Name: trimmedName})
}

// parseClassDeclaration compiles Class...End Class and registers class metadata/methods.
func (c *Compiler) parseClassDeclaration() {
	c.expectKeyword(vbscript.KeywordClass)
	className := c.expectIdentifier()
	c.registerClassDeclaration(className)
	c.predeclareClassMethodNames(className)
	classNameIdx := c.addConstant(NewString(className))
	c.emit(OpRegisterClass, classNameIdx)

	for {
		if c.matchEof() {
			panic(c.vbCompileError(vbscript.ExpectedEnd, "Expected 'End Class' before end of file"))
		}

		if c.checkKeyword(vbscript.KeywordEnd) {
			c.move()
			c.expectKeyword(vbscript.KeywordClass)
			break
		}

		if c.checkKeyword(vbscript.KeywordImplements) {
			c.parseImplementsStatement(className)
			continue
		}

		isPublic := true
		if c.checkKeyword(vbscript.KeywordPublic) {
			c.move()
			isPublic = true
		} else if c.checkKeyword(vbscript.KeywordPrivate) {
			c.move()
			isPublic = false
		}

		isDefaultMember := false
		if c.matchKeywordOrIdentifier(vbscript.KeywordDefault, "default") {
			c.move()
			isDefaultMember = true
		}

		if c.matchKeywordOrIdentifier(vbscript.KeywordSub, "sub") {
			c.parseClassMethodDeclaration(className, false, isPublic, isDefaultMember)
			continue
		}
		if c.checkKeyword(vbscript.KeywordDim) {
			if isDefaultMember {
				panic(c.vbCompileError(vbscript.ExpectedSub, "Default member must be a Sub, Function, or Property"))
			}
			c.parseClassFieldDeclaration(className, true)
			continue
		}
		if c.matchKeywordOrIdentifier(vbscript.KeywordFunction, "function") {
			c.parseClassMethodDeclaration(className, true, isPublic, isDefaultMember)
			continue
		}
		if c.matchKeywordOrIdentifier(vbscript.KeywordProperty, "property") {
			c.parseClassPropertyDeclaration(className, isPublic, isDefaultMember)
			continue
		}
		if c.matchKeywordOrIdentifier(vbscript.KeywordEvent, "event") {
			if isDefaultMember {
				panic(c.vbCompileError(vbscript.ExpectedSub, "Events cannot be the default member of a class"))
			}
			c.parseClassEventDeclaration(className)
			continue
		}
		if c.checkKeyword(vbscript.KeywordWithEvents) || (isPublic && c.isIdentifierLikeToken(c.next)) || (!isPublic && c.isIdentifierLikeToken(c.next)) {
			if isDefaultMember {
				panic(c.vbCompileError(vbscript.ExpectedSub, "Default member must be a Sub, Function, or Property"))
			}
			c.parseClassFieldDeclaration(className, isPublic)
			continue
		}
		if c.checkKeyword(vbscript.KeywordClass) {
			panic(c.vbCompileError(vbscript.SyntaxError, "Syntax error: nested Class declarations are not supported"))
		}

		c.move()
	}
}

func classDeclarationIdentifierName(tok vbscript.Token) string {
	switch t := tok.(type) {
	case *vbscript.IdentifierToken:
		return strings.TrimSpace(t.Name)
	case *vbscript.KeywordOrIdentifierToken:
		return strings.TrimSpace(t.Name)
	default:
		return ""
	}
}

// predeclareClassMethodNames scans the current class body using a cloned lexer
// cursor and pre-registers class method names so forward member calls resolve
// as class calls during compilation.
func (c *Compiler) predeclareClassMethodNames(className string) {
	if c == nil || c.lexer == nil {
		return
	}

	scan := *c.lexer
	procKind := ""

	for {
		tok := scan.NextToken()
		if tok == nil {
			return
		}
		if _, ok := tok.(*vbscript.EOFToken); ok {
			return
		}

		if procKind != "" {
			if c.tokenMatchesKeywordOrIdentifier(tok, vbscript.KeywordEnd, "end") {
				next := scan.NextToken()
				if procKind == "sub" && c.tokenMatchesKeywordOrIdentifier(next, vbscript.KeywordSub, "sub") {
					procKind = ""
					continue
				}
				if procKind == "function" && c.tokenMatchesKeywordOrIdentifier(next, vbscript.KeywordFunction, "function") {
					procKind = ""
					continue
				}
				if procKind == "property" && c.tokenMatchesKeywordOrIdentifier(next, vbscript.KeywordProperty, "property") {
					procKind = ""
					continue
				}
			}
			continue
		}

		if c.tokenMatchesKeywordOrIdentifier(tok, vbscript.KeywordEnd, "end") {
			next := scan.NextToken()
			if c.tokenMatchesKeywordOrIdentifier(next, vbscript.KeywordClass, "class") {
				return
			}
			continue
		}

		if c.tokenMatchesKeywordOrIdentifier(tok, vbscript.KeywordPublic, "public") ||
			c.tokenMatchesKeywordOrIdentifier(tok, vbscript.KeywordPrivate, "private") ||
			c.tokenMatchesKeywordOrIdentifier(tok, vbscript.KeywordDefault, "default") {
			continue
		}

		if c.tokenMatchesKeywordOrIdentifier(tok, vbscript.KeywordSub, "sub") {
			name := classDeclarationIdentifierName(scan.NextToken())
			if name != "" {
				c.addClassMethodDeclaration(className, CompiledClassMethodDecl{Name: name, IsFunction: false})
			}
			procKind = "sub"
			continue
		}

		if c.tokenMatchesKeywordOrIdentifier(tok, vbscript.KeywordFunction, "function") {
			name := classDeclarationIdentifierName(scan.NextToken())
			if name != "" {
				c.addClassMethodDeclaration(className, CompiledClassMethodDecl{Name: name, IsFunction: true})
			}
			procKind = "function"
			continue
		}

		if c.tokenMatchesKeywordOrIdentifier(tok, vbscript.KeywordProperty, "property") {
			accessor := scan.NextToken()
			accessorStr := ""
			if c.tokenMatchesKeywordOrIdentifier(accessor, vbscript.KeywordGet, "get") {
				accessorStr = "get"
			} else if c.tokenMatchesKeywordOrIdentifier(accessor, vbscript.KeywordLet, "let") {
				accessorStr = "let"
			} else if c.tokenMatchesKeywordOrIdentifier(accessor, vbscript.KeywordSet, "set") {
				accessorStr = "set"
			}
			name := classDeclarationIdentifierName(scan.NextToken())
			if name != "" && accessorStr != "" {
				paramCount := preScanCountPropertyParams(&scan)
				c.preDeclareClassPropertyAccessor(className, name, accessorStr, paramCount)
			}
			procKind = "property"
			continue
		}
	}
}

// preScanCountPropertyParams counts Property accessor parameters from a pre-scan cursor.
func preScanCountPropertyParams(scan *vbscript.Lexer) int {
	next := scan.NextToken()
	lp, ok := next.(*vbscript.PunctuationToken)
	if !ok || lp.Type != vbscript.PunctLParen {
		return 0
	}
	commas := 0
	hasContent := false
	depth := 1
	for depth > 0 {
		tok := scan.NextToken()
		if tok == nil {
			break
		}
		if _, isEOF := tok.(*vbscript.EOFToken); isEOF {
			break
		}
		if p, ok2 := tok.(*vbscript.PunctuationToken); ok2 {
			switch p.Type {
			case vbscript.PunctLParen:
				depth++
			case vbscript.PunctRParen:
				depth--
			case vbscript.PunctComma:
				if depth == 1 {
					commas++
				}
			}
			continue
		}
		if depth == 1 {
			switch tok.(type) {
			case *vbscript.LineTerminationToken, *vbscript.ColonLineTerminationToken, *vbscript.CommentToken:
			default:
				hasContent = true
			}
		}
	}
	if commas > 0 {
		return commas + 1
	}
	if hasContent {
		return 1
	}
	return 0
}

// preDeclareClassPropertyAccessor pre-registers one Property accessor from the pre-scan pass.
func (c *Compiler) preDeclareClassPropertyAccessor(className, propertyName, accessor string, paramCount int) {
	existing, exists := c.getClassPropertyDeclaration(className, propertyName)
	var prop CompiledClassPropertyDecl
	if exists && existing != nil {
		prop = *existing
	} else {
		prop = CompiledClassPropertyDecl{
			Name:              propertyName,
			GetUserSubConstID: -1,
			LetUserSubConstID: -1,
			SetUserSubConstID: -1,
		}
	}
	switch accessor {
	case "get":
		if !prop.HasGet {
			prop.HasGet = true
			prop.GetParamCount = paramCount
			prop.GetPreScanned = true
		}
	case "let":
		if !prop.HasLet {
			prop.HasLet = true
			prop.LetParamCount = paramCount
			prop.LetPreScanned = true
		}
	case "set":
		if !prop.HasSet {
			prop.HasSet = true
			prop.SetParamCount = paramCount
			prop.SetPreScanned = true
		}
	}
	c.setClassPropertyDeclaration(className, prop)
}

// parseClassFieldDeclaration registers one or more class fields without emitting runtime top-level code.
func (c *Compiler) parseClassFieldDeclaration(className string, isPublic bool) {
	if c.checkKeyword(vbscript.KeywordDim) {
		c.move()
	}

	withEvents := false
	if c.checkKeyword(vbscript.KeywordWithEvents) {
		c.move()
		withEvents = true
	}

	for {
		fieldName := c.expectIdentifier()
		c.addClassFieldDeclaration(className, CompiledClassFieldDecl{Name: fieldName, IsPublic: isPublic, WithEvents: withEvents})

		classNameIdx := c.addConstant(NewString(className))
		fieldNameIdx := c.addConstant(NewString(fieldName))
		isPublicOperand := 0
		if isPublic {
			isPublicOperand = 1
		}
		c.emit(OpRegisterClassField, classNameIdx, fieldNameIdx, isPublicOperand)

		if withEvents {
			c.emitExt(ExtOpWithEventsRegister, classNameIdx, fieldNameIdx)
		}

		// If the field has array dimensions (e.g., Private arr(5) or Private arr(3,5)),
		// parse each upper-bound expression and emit OpInitClassArrayField so the VM can
		// allocate a pre-sized array for every new class instance.
		if p, ok := c.next.(*vbscript.PunctuationToken); ok && p.Type == vbscript.PunctLParen {
			c.move() // consume '('
			valCount := 0
			// Empty parens Private arr() means dynamic (undimensioned); emit no init opcode.
			if rp, ok2 := c.next.(*vbscript.PunctuationToken); !(ok2 && rp.Type == vbscript.PunctRParen) {
				for {
					c.parseExpression(PrecNone) // pushes upper-bound value onto the stack
					if c.checkKeyword(vbscript.KeywordTo) {
						c.move()
						c.parseExpression(PrecNone)
						valCount += 2
					} else {
						// Push default lower bound and swap
						c.emit(OpConstant, c.addConstant(NewInteger(int64(c.optionBase))))
						c.emit(OpSwap)
						valCount += 2
					}
					if comma, ok3 := c.next.(*vbscript.PunctuationToken); ok3 && comma.Type == vbscript.PunctComma {
						c.move()
						continue
					}
					break
				}
			}
			if rp, ok2 := c.next.(*vbscript.PunctuationToken); !ok2 || rp.Type != vbscript.PunctRParen {
				panic(c.vbCompileError(vbscript.ExpectedRParen, "Expected ')' after array bounds"))
			}
			c.move() // consume ')'
			if valCount > 0 {
				c.emit(OpInitClassArrayField, classNameIdx, fieldNameIdx, valCount)
			}
		}

		if p, ok := c.next.(*vbscript.PunctuationToken); ok && p.Type == vbscript.PunctComma {
			c.move()
			continue
		}
		break
	}
	for {
		switch c.next.(type) {
		case *vbscript.LineTerminationToken, *vbscript.ColonLineTerminationToken, *vbscript.CommentToken:
			return
		}
		break
	}
}

// tokenMatchesKeywordOrIdentifier reports whether one arbitrary token matches one keyword or keyword-like text.
func (c *Compiler) tokenMatchesKeywordOrIdentifier(token vbscript.Token, kw vbscript.Keyword, text string) bool {
	if token == nil {
		return false
	}
	switch t := token.(type) {
	case *vbscript.KeywordToken:
		return t.Keyword == kw
	case *vbscript.KeywordOrIdentifierToken:
		return strings.EqualFold(strings.TrimSpace(t.Name), text)
	case *vbscript.IdentifierToken:
		return strings.EqualFold(strings.TrimSpace(t.Name), text)
	default:
		return false
	}
}

// procParamResult holds the result of parsing procedure formal parameters.
type procParamResult struct {
	names         []string
	byRefMask     uint64
	optionalMask  uint64
	paramArrayIdx int   // index of ParamArray param, -1 if none
	defaults      []int // constant pool indices for default values, -1 for params without defaults
}

// parseProcedureParameterNames parses Sub/Function formal parameter names and modifiers.
// Supports VB6 advanced signatures: ByRef, ByVal, Optional [As Type] = DefaultValue, ParamArray.
// Returns the parameter metadata including names, masks, default value constant indices.
func (c *Compiler) parseProcedureParameterNames() procParamResult {
	result := procParamResult{
		paramArrayIdx: -1,
	}
	if p, ok := c.next.(*vbscript.PunctuationToken); !ok || p.Type != vbscript.PunctLParen {
		return result
	}

	c.move()
	paramIdx := 0
	for {
		if p, ok := c.next.(*vbscript.PunctuationToken); ok && p.Type == vbscript.PunctRParen {
			break
		}
		if c.matchEof() {
			break
		}

		// VBScript default for parameters without an explicit modifier is ByRef.
		isByRef := true
		isOptional := false
		isParamArray := false

		// Parse ByRef/ByVal/Optional/ParamArray modifiers in any order.
	parsedModifiers:
		for {
			if c.checkKeyword(vbscript.KeywordByRef) {
				c.move()
				isByRef = true
			} else if c.checkKeyword(vbscript.KeywordByVal) {
				c.move()
				isByRef = false
			} else if c.matchKeywordOrIdentifier(vbscript.KeywordOptional, "optional") {
				c.move()
				isOptional = true
			} else if c.matchKeywordOrIdentifier(vbscript.KeywordParamArray, "paramarray") {
				c.move()
				isParamArray = true
				isByRef = false // ParamArray is always ByVal
				result.paramArrayIdx = paramIdx
			} else {
				break parsedModifiers
			}
		}

		paramName := c.expectIdentifier()
		result.names = append(result.names, paramName)

		// For ParamArray parameters, consume the trailing "()" syntax.
		// VB6 syntax: ParamArray values()  -- the () marks it as an array.
		if isParamArray {
			if p, ok := c.next.(*vbscript.PunctuationToken); ok && p.Type == vbscript.PunctLParen {
				c.move() // consume '('
				if p2, ok2 := c.next.(*vbscript.PunctuationToken); ok2 && p2.Type == vbscript.PunctRParen {
					c.move() // consume ')'
				} else {
					// If there's something between the parens, this is likely an error.
					panic(c.vbCompileError(vbscript.SyntaxError,
						"ParamArray parameter must be declared with empty parentheses: values()"))
				}
			} else {
				panic(c.vbCompileError(vbscript.SyntaxError,
					"ParamArray parameter requires empty parentheses: values()"))
			}
		}

		// Parse optional As Type clause.
		declaredType, udtName := c.parseAsTypeClause()
		if declaredType != VTEmpty {
			// Record type declaration for the parameter in the local scope.
			lower := strings.ToLower(paramName)
			c.localVarTypes[lower] = declaredType
			if declaredType == VTRecord {
				c.localRecordTypes[lower] = udtName
			}
		}

		// Parse optional = DefaultValue for Optional parameters.
		defaultIdx := -1
		if isOptional {
			if p, ok := c.next.(*vbscript.PunctuationToken); ok && p.Type == vbscript.PunctEqual {
				c.move()
				// Compile the default value expression.
				c.parseExpression(PrecNone)
				defaultIdx = c.extractDefaultConst()
			}
		}

		result.defaults = append(result.defaults, defaultIdx)

		// Set byRefMask: only ByRef params get a bit set.
		if isByRef && !isParamArray && paramIdx < 64 {
			result.byRefMask |= 1 << uint(paramIdx)
		}

		// Set optionalMask.
		if isOptional && paramIdx < 64 {
			result.optionalMask |= 1 << uint(paramIdx)
		}

		// Validate ParamArray: must be the last parameter.
		if isParamArray {
			// Check there's no comma after this parameter (it must be last).
			if p, ok := c.next.(*vbscript.PunctuationToken); ok && p.Type == vbscript.PunctComma {
				panic(c.vbCompileError(vbscript.SyntaxError, "ParamArray must be the last parameter"))
			}
		}

		paramIdx++
		if p, ok := c.next.(*vbscript.PunctuationToken); ok && p.Type == vbscript.PunctComma {
			c.move()
		}
	}
	if p, ok := c.next.(*vbscript.PunctuationToken); ok && p.Type == vbscript.PunctRParen {
		c.move()
	}

	return result
}

// extractDefaultConst extracts a constant index from the most recently compiled expression
// (used for Optional parameter default values). It scans backwards from the end of bytecode
// for the trailing OpConstant followed by a 2-byte constant index, skipping any OpNop
// placeholders left by the constant folder. If the expression is not a simple compile-time
// constant (e.g. a function call), it raises a compile error since VB6 requires Optional
// parameter defaults to be constant expressions.
func (c *Compiler) extractDefaultConst() int {
	bc := c.bytecode
	// Scan backwards, skipping OpNop placeholders.
	i := len(bc)
	for i >= 3 && OpCode(bc[i-3]) == OpNop {
		i -= 3
	}
	if i >= 3 && OpCode(bc[i-3]) == OpConstant {
		idx := int(binary.BigEndian.Uint16(bc[i-2:]))
		// Remove the trailing opcode(s) from bytecode.
		c.bytecode = bc[:i-3]
		return idx
	}
	// The expression is not a simple compile-time constant.
	panic(c.vbCompileError(vbscript.SyntaxError,
		"Optional parameter default value must be a constant expression"))
}

// parseClassMethodDeclaration compiles one class Sub/Function body and registers runtime class method metadata.
func (c *Compiler) parseClassMethodDeclaration(className string, isFunc bool, isPublic bool, isDefaultMember bool) {
	if isFunc {
		if c.matchKeywordOrIdentifier(vbscript.KeywordFunction, "function") {
			c.move()
		} else {
			c.expectKeyword(vbscript.KeywordFunction)
		}
	} else {
		if c.matchKeywordOrIdentifier(vbscript.KeywordSub, "sub") {
			c.move()
		} else {
			c.expectKeyword(vbscript.KeywordSub)
		}
	}

	methodName := c.expectIdentifier()
	paramResult := c.parseProcedureParameterNames()
	if strings.EqualFold(methodName, "Class_Initialize") || strings.EqualFold(methodName, "Class_Terminate") {
		if len(paramResult.names) != 0 {
			panic(c.vbCompileError(vbscript.ClassInitializeOrTerminateDoNotHaveArguments, "Class_Initialize/Class_Terminate must not declare arguments"))
		}
	}

	placeholder := c.addConstant(NewEmpty())
	jumpOverBody := c.emitJump(OpJump)

	entryPoint := len(c.bytecode)
	c.constants[placeholder] = NewUserSubEx(entryPoint, len(paramResult.names), c.locals.Count(), isFunc, paramResult.byRefMask, paramResult.optionalMask, paramResult.paramArrayIdx, nil)

	// Store default value constant indices for Optional parameters.
	if len(paramResult.defaults) > 0 {
		defaults := make([]int, len(paramResult.defaults))
		copy(defaults, paramResult.defaults)
		c.funcParamDefaults[entryPoint] = defaults
	}

	prevIsLocal := c.isLocal
	prevLocals := c.locals
	prevDeclared := c.declaredLocals
	prevConstLocals := c.constLocals
	prevStaticLocals := c.staticLocals
	prevClassName := c.currentClassName
	prevFunctionName := c.currentFunctionName
	prevLabelMap := c.labelMap
	prevForwardLabelPatches := c.forwardLabelPatches

	c.isLocal = true
	c.currentClassName = className
	c.locals = NewSymbolTable()
	c.declaredLocals = make(map[string]bool)
	c.constLocals = make(map[string]bool)
	c.staticLocals = make(map[string]int)
	c.labelMap = make(map[string]int)
	c.forwardLabelPatches = make(map[string][]int)
	c.currentFunctionName = methodName

	for _, p := range paramResult.names {
		c.locals.Add(p)
		c.declaredLocals[strings.ToLower(p)] = true
	}

	returnIdx := -1
	if isFunc {
		returnIdx = c.locals.Add(methodName)
		c.declaredLocals[strings.ToLower(methodName)] = true
	}

	c.hoistProcedureDimDeclarations(keywordFromBool(isFunc))

	for !c.matchEof() {
		if c.checkKeyword(vbscript.KeywordEnd) {
			peek := c.peekToken()
			if isFunc {
				if c.tokenMatchesKeywordOrIdentifier(peek, vbscript.KeywordFunction, "function") {
					break
				}
			} else {
				if c.tokenMatchesKeywordOrIdentifier(peek, vbscript.KeywordSub, "sub") {
					break
				}
			}
		}
		if c.matchEof() {
			break
		}
		c.parseStatement()
	}

	c.expectKeyword(vbscript.KeywordEnd)
	if isFunc {
		if c.matchKeywordOrIdentifier(vbscript.KeywordFunction, "function") {
			c.move()
		} else {
			c.expectKeyword(vbscript.KeywordFunction)
		}
	} else {
		if c.matchKeywordOrIdentifier(vbscript.KeywordSub, "sub") {
			c.move()
		} else {
			c.expectKeyword(vbscript.KeywordSub)
		}
	}

	c.constants[placeholder] = NewUserSubEx(entryPoint, len(paramResult.names), c.locals.Count(), isFunc, paramResult.byRefMask, paramResult.optionalMask, paramResult.paramArrayIdx, c.locals.names)

	if isFunc {
		c.emit(OpGetLocal, returnIdx)
	} else {
		emptyIdx := c.addConstant(NewEmpty())
		c.emit(OpConstant, emptyIdx)
	}
	c.emit(OpRet)

	c.patchJump(jumpOverBody)

	classNameIdx := c.addConstant(NewString(className))
	methodNameIdx := c.addConstant(NewString(methodName))
	isPublicOperand := 0
	if isPublic {
		isPublicOperand = 1
	}
	c.emit(OpRegisterClassMethod, classNameIdx, methodNameIdx, placeholder, isPublicOperand)
	if isDefaultMember {
		defaultNameIdx := c.addConstant(NewString("__default__"))
		c.emit(OpRegisterClassMethod, classNameIdx, defaultNameIdx, placeholder, isPublicOperand)
	}

	c.addClassMethodDeclaration(className, CompiledClassMethodDecl{
		Name:           methodName,
		IsFunction:     isFunc,
		IsPublic:       isPublic,
		UserSubConstID: placeholder,
	})

	if len(c.forwardLabelPatches) > 0 {
		for label := range c.forwardLabelPatches {
			panic(c.vbCompileError(vbscript.SyntaxError, fmt.Sprintf("Label '%s' not defined in method '%s.%s'", label, className, methodName)))
		}
	}

	c.isLocal = prevIsLocal
	c.currentClassName = prevClassName
	c.locals = prevLocals
	c.declaredLocals = prevDeclared
	c.constLocals = prevConstLocals
	c.staticLocals = prevStaticLocals
	c.currentFunctionName = prevFunctionName
	c.labelMap = prevLabelMap
	c.forwardLabelPatches = prevForwardLabelPatches
}

// parseClassPropertyDeclaration compiles one class Property Get/Let/Set body and validates signatures.
func (c *Compiler) parseClassPropertyDeclaration(className string, isPublic bool, isDefaultMember bool) {
	if c.matchKeywordOrIdentifier(vbscript.KeywordProperty, "property") {
		c.move()
	} else {
		panic(c.vbCompileError(vbscript.ExpectedProperty, "Expected Property declaration"))
	}

	accessorKind := classPropertyAccessorGet
	isFunction := false
	if c.matchKeywordOrIdentifier(vbscript.KeywordGet, "get") {
		c.move()
		accessorKind = classPropertyAccessorGet
		isFunction = true
	} else if c.matchKeywordOrIdentifier(vbscript.KeywordLet, "let") {
		c.move()
		accessorKind = classPropertyAccessorLet
	} else if c.matchKeywordOrIdentifier(vbscript.KeywordSet, "set") {
		c.move()
		accessorKind = classPropertyAccessorSet
	} else {
		panic(c.vbCompileError(vbscript.ExpectedLetGetSet, "Expected Property Get, Let, or Set declaration"))
	}

	propertyName := c.expectIdentifier()
	paramResult := c.parseProcedureParameterNames()

	if (accessorKind == classPropertyAccessorLet || accessorKind == classPropertyAccessorSet) && len(paramResult.names) == 0 {
		panic(c.vbCompileError(vbscript.PropertySetOrLetMustHaveArguments, "Property Let/Set requires one value parameter"))
	}

	placeholder := c.addConstant(NewEmpty())
	jumpOverBody := c.emitJump(OpJump)

	entryPoint := len(c.bytecode)
	c.constants[placeholder] = NewUserSubEx(entryPoint, len(paramResult.names), c.locals.Count(), isFunction, paramResult.byRefMask, paramResult.optionalMask, paramResult.paramArrayIdx, nil)

	// Store default value constant indices for Optional parameters.
	if len(paramResult.defaults) > 0 {
		defaults := make([]int, len(paramResult.defaults))
		copy(defaults, paramResult.defaults)
		c.funcParamDefaults[entryPoint] = defaults
	}

	prevIsLocal := c.isLocal
	prevLocals := c.locals
	prevDeclared := c.declaredLocals
	prevConstLocals := c.constLocals
	prevStaticLocals := c.staticLocals
	prevClassName := c.currentClassName
	prevFunctionName := c.currentFunctionName
	prevLabelMap := c.labelMap
	prevForwardLabelPatches := c.forwardLabelPatches

	c.isLocal = true
	c.currentClassName = className
	c.locals = NewSymbolTable()
	c.declaredLocals = make(map[string]bool)
	c.constLocals = make(map[string]bool)
	c.staticLocals = make(map[string]int)
	c.labelMap = make(map[string]int)
	c.forwardLabelPatches = make(map[string][]int)
	c.currentFunctionName = propertyName

	for _, p := range paramResult.names {
		c.locals.Add(p)
		c.declaredLocals[strings.ToLower(p)] = true
	}

	returnIdx := -1
	if isFunction {
		returnIdx = c.locals.Add(propertyName)
		c.declaredLocals[strings.ToLower(propertyName)] = true
	}

	c.hoistProcedureDimDeclarations(vbscript.KeywordProperty)

	for !c.matchEof() {
		if c.checkKeyword(vbscript.KeywordEnd) {
			peek := c.peekToken()
			if c.tokenMatchesKeywordOrIdentifier(peek, vbscript.KeywordProperty, "property") {
				break
			}
		}
		if c.matchEof() {
			break
		}
		c.parseStatement()
	}

	c.expectKeyword(vbscript.KeywordEnd)
	if c.matchKeywordOrIdentifier(vbscript.KeywordProperty, "property") {
		c.move()
	} else {
		c.expectKeyword(vbscript.KeywordProperty)
	}

	c.constants[placeholder] = NewUserSubEx(entryPoint, len(paramResult.names), c.locals.Count(), isFunction, paramResult.byRefMask, paramResult.optionalMask, paramResult.paramArrayIdx, c.locals.names)

	if isFunction {
		c.emit(OpGetLocal, returnIdx)
	} else {
		emptyIdx := c.addConstant(NewEmpty())
		c.emit(OpConstant, emptyIdx)
	}
	c.emit(OpRet)

	c.patchJump(jumpOverBody)

	if len(c.forwardLabelPatches) > 0 {
		for label := range c.forwardLabelPatches {
			panic(c.vbCompileError(vbscript.SyntaxError, fmt.Sprintf("Label '%s' not defined in property '%s.%s'", label, className, propertyName)))
		}
	}

	c.registerClassPropertyDeclaration(className, propertyName, isPublic, accessorKind, len(paramResult.names), placeholder)

	classNameIdx := c.addConstant(NewString(className))
	propertyNameIdx := c.addConstant(NewString(propertyName))
	isPublicOperand := 0
	if isPublic {
		isPublicOperand = 1
	}
	switch accessorKind {
	case classPropertyAccessorGet:
		c.emit(OpRegisterClassPropertyGet, classNameIdx, propertyNameIdx, placeholder, len(paramResult.names), isPublicOperand)
	case classPropertyAccessorLet:
		c.emit(OpRegisterClassPropertyLet, classNameIdx, propertyNameIdx, placeholder, len(paramResult.names), isPublicOperand)
	case classPropertyAccessorSet:
		c.emit(OpRegisterClassPropertySet, classNameIdx, propertyNameIdx, placeholder, len(paramResult.names), isPublicOperand)
	}

	if isDefaultMember {
		defaultNameIdx := c.addConstant(NewString("__default__"))
		switch accessorKind {
		case classPropertyAccessorGet:
			c.emit(OpRegisterClassPropertyGet, classNameIdx, defaultNameIdx, placeholder, len(paramResult.names), isPublicOperand)
		case classPropertyAccessorLet:
			c.emit(OpRegisterClassPropertyLet, classNameIdx, defaultNameIdx, placeholder, len(paramResult.names), isPublicOperand)
		case classPropertyAccessorSet:
			c.emit(OpRegisterClassPropertySet, classNameIdx, defaultNameIdx, placeholder, len(paramResult.names), isPublicOperand)
		}
	}

	c.isLocal = prevIsLocal
	c.currentClassName = prevClassName
	c.locals = prevLocals
	c.declaredLocals = prevDeclared
	c.constLocals = prevConstLocals
	c.staticLocals = prevStaticLocals
	c.currentFunctionName = prevFunctionName
	c.labelMap = prevLabelMap
	c.forwardLabelPatches = prevForwardLabelPatches
}

// registerClassPropertyDeclaration validates and stores one Property accessor metadata entry.
func (c *Compiler) registerClassPropertyDeclaration(className string, propertyName string, isPublic bool, accessorKind classPropertyAccessorKind, paramCount int, userSubConstID int) {
	property, exists := c.getClassPropertyDeclaration(className, propertyName)
	if !exists {
		newProperty := CompiledClassPropertyDecl{
			Name:              propertyName,
			IsPublic:          isPublic,
			GetUserSubConstID: -1,
			LetUserSubConstID: -1,
			SetUserSubConstID: -1,
		}
		property = &newProperty
	}

	property.IsPublic = isPublic

	signatureError := func() {
		panic(c.vbCompileError(vbscript.InconsistentNumberOfArguments, "Property signature mismatch between Get/Let/Set accessors"))
	}

	switch accessorKind {
	case classPropertyAccessorGet:
		if property.HasGet && !property.GetPreScanned {
			panic(c.vbCompileError(vbscript.SyntaxError, "Property Get already defined"))
		}
		property.GetPreScanned = false
		if property.HasLet && !property.LetPreScanned && property.LetParamCount != paramCount+1 {
			signatureError()
		}
		if property.HasSet && !property.SetPreScanned && property.SetParamCount != paramCount+1 {
			signatureError()
		}
		property.HasGet = true
		property.GetParamCount = paramCount
		property.GetUserSubConstID = userSubConstID
	case classPropertyAccessorLet:
		if property.HasLet && !property.LetPreScanned {
			panic(c.vbCompileError(vbscript.SyntaxError, "Property Let already defined"))
		}
		property.LetPreScanned = false
		if paramCount < 1 {
			signatureError()
		}
		if property.HasGet && !property.GetPreScanned && property.GetParamCount != paramCount-1 {
			signatureError()
		}
		if property.HasSet && !property.SetPreScanned && property.SetParamCount != paramCount {
			signatureError()
		}
		property.HasLet = true
		property.LetParamCount = paramCount
		property.LetUserSubConstID = userSubConstID
	case classPropertyAccessorSet:
		if property.HasSet && !property.SetPreScanned {
			panic(c.vbCompileError(vbscript.SyntaxError, "Property Set already defined"))
		}
		property.SetPreScanned = false
		if paramCount < 1 {
			signatureError()
		}
		if property.HasGet && !property.GetPreScanned && property.GetParamCount != paramCount-1 {
			signatureError()
		}
		if property.HasLet && !property.LetPreScanned && property.LetParamCount != paramCount {
			signatureError()
		}
		property.HasSet = true
		property.SetParamCount = paramCount
		property.SetUserSubConstID = userSubConstID
	}

	c.setClassPropertyDeclaration(className, *property)
}

// parseParenArgumentList parses a parenthesized comma-separated expression list and returns argument count.
func (c *Compiler) parseParenArgumentList() int {
	openParen, ok := c.next.(*vbscript.PunctuationToken)
	if !ok || openParen.Type != vbscript.PunctLParen {
		panic(c.vbCompileError(vbscript.ExpectedLParen, "Expected '(' before argument list"))
	}
	c.move() // Consume '('

	argCount := 0
	if closeParen, ok := c.next.(*vbscript.PunctuationToken); !ok || closeParen.Type != vbscript.PunctRParen {
		for {
			if comma, ok := c.next.(*vbscript.PunctuationToken); ok && comma.Type == vbscript.PunctComma {
				emptyIdx := c.addConstant(NewEmpty())
				c.emit(OpConstant, emptyIdx)
			} else if closeParen, ok := c.next.(*vbscript.PunctuationToken); ok && closeParen.Type == vbscript.PunctRParen {
				emptyIdx := c.addConstant(NewEmpty())
				c.emit(OpConstant, emptyIdx)
			} else {
				pargStartPos := len(c.bytecode)
				c.parseExpression(PrecNone)
				c.patchArgRefInBytecode(pargStartPos)
			}
			argCount++
			if comma, ok := c.next.(*vbscript.PunctuationToken); ok && comma.Type == vbscript.PunctComma {
				c.move()
				continue
			}
			break
		}
	}

	closeParen, ok := c.next.(*vbscript.PunctuationToken)
	if !ok || closeParen.Type != vbscript.PunctRParen {
		panic(c.vbCompileError(vbscript.ExpectedRParen, "Expected ')' after argument list"))
	}
	c.move() // Consume ')'
	return argCount
}

// parseStaticStatement compiles Static local variable declarations as mapped global slots.
func (c *Compiler) parseStaticStatement() {
	c.expectKeyword(vbscript.KeywordStatic)
	for {
		name := c.expectIdentifier()
		lower := strings.ToLower(name)

		// Static variables should have been hoisted into staticLocals.
		globalIdx, ok := c.staticLocals[lower]
		if !ok {
			// Fallback if hoisting missed it
			prefix := ""
			if c.currentClassName != "" {
				prefix = c.currentClassName + "_"
			}
			hiddenName := fmt.Sprintf("__static_%s%s_%s", prefix, c.currentFunctionName, name)
			globalIdx = c.Globals.Add(hiddenName)
			c.staticLocals[lower] = globalIdx
			c.declaredGlobals[strings.ToLower(hiddenName)] = true
		}

		declaredType, udtName := c.parseAsTypeClause()

		// If it has array bounds OR As Type, we need a guard
		isArray := false
		if p, ok := c.next.(*vbscript.PunctuationToken); ok && p.Type == vbscript.PunctLParen {
			isArray = true
		}

		if isArray || declaredType != VTEmpty {
			// Guard: If IsEmpty(__static_...) Then
			c.emitBuiltinTarget("IsEmpty")
			c.emit(OpGetGlobal, globalIdx)
			c.emit(OpCall, 1)
			jumpIfNotEmpty := c.emitJump(OpJumpIfFalse)

			if isArray {
				c.tryParseArrayDeclaration(name)
			}

			// For non-array, non-record types, initialize to default
			if declaredType != VTEmpty && declaredType != VTRecord && !isArray {
				var val Value
				switch declaredType {
				case VTInteger:
					val = NewInteger(0)
				case VTDouble:
					val = NewDouble(0)
				case VTString:
					val = NewString("")
				case VTBool:
					val = NewBool(false)
				case VTDate:
					val = NewDate(time.Time{})
				default:
					val = NewEmpty()
				}
				c.emit(OpConstant, c.addConstant(val))
				opSet, idxSet := c.resolveSetVar(name)
				c.emit(c.letOpCode(opSet), idxSet)
			}

			if declaredType == VTRecord {
				// emitTypedInit will handle the initialization
				c.emitTypedInit(name, declaredType, udtName)
			}

			c.patchJump(jumpIfNotEmpty)

			// Still call emitTypedInit to register the type mappings (without re-initializing record)
			if declaredType != VTRecord {
				c.emitTypedInit(name, declaredType, udtName)
			}
		}

		if comma, ok := c.next.(*vbscript.PunctuationToken); ok && comma.Type == vbscript.PunctComma {
			c.move()
			continue
		}
		break
	}
	c.skipStatementEnd()
}

// parseConstStatement compiles Const declarations as declared variables initialized from constant expressions.
func (c *Compiler) parseConstStatement() {
	c.expectKeyword(vbscript.KeywordConst)
	for {
		name := c.expectIdentifier()
		c.declareConst(name)

		if p, ok := c.next.(*vbscript.PunctuationToken); !ok || p.Type != vbscript.PunctEqual {
			panic(c.vbCompileError(vbscript.ExpectedEqual, "Expected '=' in Const declaration"))
		}
		c.move()

		exprStart := len(c.bytecode)
		c.parseExpression(PrecNone)
		exprEnd := len(c.bytecode)
		if !c.isLocal && c.currentClassName == "" {
			c.tryCaptureGlobalConstLiteral(name, exprStart, exprEnd)
			if literalValue, ok := c.constLiteralGlobals[strings.ToLower(strings.TrimSpace(name))]; ok {
				c.patchForwardConstSites(name, literalValue)
			}
		}
		op, idx := c.resolveConstInitVar(name)
		c.emit(op, idx)

		if p, ok := c.next.(*vbscript.PunctuationToken); ok && p.Type == vbscript.PunctComma {
			c.move()
			continue
		}
		break
	}
}

// parseEnumStatement compiles Enum declarations as compile-time constants.
func (c *Compiler) parseEnumStatement() {
	c.expectKeyword(vbscript.KeywordEnum)
	_ = c.expectIdentifier() // Enum name (ignored at runtime)
	c.skipStatementEnd()

	var currentValue int64 = 0

	for {
		if c.matchKeywordOrIdentifier(vbscript.KeywordEnd, "end") {
			c.move()
			c.expectKeyword(vbscript.KeywordEnum)
			c.skipStatementEnd()
			break
		}

		if c.matchEof() {
			panic(c.vbCompileError(vbscript.ExpectedEnd, "Expected 'End Enum'"))
		}

		if c.isStatementEnd() {
			c.skipStatementEnd()
			continue
		}

		name := c.expectIdentifier()
		lower := strings.ToLower(name)

		if p, ok := c.next.(*vbscript.PunctuationToken); ok && p.Type == vbscript.PunctEqual {
			c.move()
			// For simplicity in single-pass, we only support integer literals or existing constants in Enum.
			if lit, ok := c.next.(*vbscript.DecIntegerLiteralToken); ok {
				currentValue = lit.Value
				c.move()
			} else if lit, ok := c.next.(*vbscript.HexIntegerLiteralToken); ok {
				currentValue = lit.Value
				c.move()
			} else {
				// Try to resolve as a constant
				constName := c.expectIdentifier()
				if val, ok := c.constLiteralGlobals[strings.ToLower(constName)]; ok && val.Type == VTInteger {
					currentValue = val.Num
				} else {
					panic(c.vbCompileError(vbscript.ExpectedConstantExpression, "Expected constant integer expression in Enum"))
				}
			}
		}

		c.constGlobals[lower] = true
		c.constLiteralGlobals[lower] = NewInteger(currentValue)
		currentValue++

		c.skipStatementEnd()
	}
}

func (c *Compiler) isStatementEnd() bool {
	if c.matchEof() {
		return true
	}
	switch c.next.(type) {
	case *vbscript.LineTerminationToken, *vbscript.ColonLineTerminationToken, *vbscript.CommentToken, *vbscript.ASPCodeEndToken:
		return true
	}
	return false
}

func (c *Compiler) skipStatementEnd() {
	for !c.matchEof() && c.isStatementEnd() {
		c.move()
	}
}

// parseScopedVariableDeclaration compiles page-scope Public/Private variable declarations.
// Classic ASP accepts module-level declarations like `Private counter` as aliases for
// page-level variable declarations; they should not be rejected as missing procedures.
func (c *Compiler) parseScopedVariableDeclaration() {
	withEvents := false
	if c.checkKeyword(vbscript.KeywordWithEvents) {
		c.move()
		withEvents = true
	}

	for {
		name := c.expectIdentifier()
		c.declareVar(name)

		// Parse optional VB6 As Type clause.
		declaredType, udtName := c.parseAsTypeClause()

		if c.tryParseArrayDeclaration(name) {
			if p, ok := c.next.(*vbscript.PunctuationToken); ok && p.Type == vbscript.PunctEqual {
				panic(c.vbSyntaxError(vbscript.ExpectedEndOfStatement))
			}
		} else {
			if p, ok := c.next.(*vbscript.PunctuationToken); ok && p.Type == vbscript.PunctEqual {
				panic(c.vbSyntaxError(vbscript.ExpectedEndOfStatement))
			}
		}

		// Emit type initialization opcode if As Type was specified.
		c.emitTypedInit(name, declaredType, udtName)

		if withEvents {
			c.emitExt(ExtOpWithEventsRegister, 0xFFFF, c.addConstant(NewString(name)))
		}

		if p, ok := c.next.(*vbscript.PunctuationToken); ok && p.Type == vbscript.PunctComma {
			c.move()
			continue
		}
		break
	}
}

// vb6TypeNameToValueType converts a VB6 type name string to the corresponding ValueType.
// Returns VTEmpty for "Variant" (meaning no type constraint, standard VBScript behavior).
// Returns (VTEmpty, false) if the type name is not recognized.
func vb6TypeNameToValueType(name string) (ValueType, bool) {
	switch strings.ToLower(strings.TrimSpace(name)) {
	case "integer":
		return VTInteger, true
	case "long":
		return VTInteger, true
	case "single":
		return VTDouble, true
	case "double":
		return VTDouble, true
	case "string":
		return VTString, true
	case "boolean":
		return VTBool, true
	case "byte":
		return VTInteger, true
	case "object":
		return VTObject, true
	case "variant":
		return VTEmpty, true // Variant = no constraint (standard behavior)
	default:
		return VTEmpty, false
	}
}

// matchAsKeyword checks whether the next token is the "As" keyword (case-insensitive).
func (c *Compiler) matchAsKeyword() bool {
	return c.matchKeywordOrIdentifier(vbscript.KeywordAs, "as")
}

// parseAsTypeClause checks for an optional "As Type" clause and returns the declared type and UDT/Class name if applicable.
func (c *Compiler) parseAsTypeClause() (ValueType, string) {
	if !c.matchAsKeyword() {
		return VTEmpty, "" // No type declared = Variant
	}
	c.move() // consume "As"
	typeName := c.expectIdentifier()
	declaredType, ok := vb6TypeNameToValueType(typeName)
	if ok {
		return declaredType, ""
	}
	// Check if it's a UDT
	if _, exists := c.recordDeclLookup[strings.ToLower(typeName)]; exists {
		return VTRecord, typeName
	}
	// Phase 5: Support Classes/Interfaces in As clause.
	// Since we are single-pass, we might not have all classes declared yet.
	// We'll treat any other identifier as an Object type (Class reference).
	return VTObject, typeName
}

// emitTypedInit records a VB6 As Type declaration for a variable in the compiler's type maps.
func (c *Compiler) emitTypedInit(name string, declaredType ValueType, udtName string) {
	if declaredType == VTEmpty {
		return // No type declaration, standard variant behavior
	}
	lower := strings.ToLower(name)

	isStatic := false
	var globalName string
	if c.isLocal {
		if globalIdx, ok := c.staticLocals[lower]; ok {
			isStatic = true
			globalName = strings.ToLower(c.Globals.names[globalIdx])
		}
	}

	if isStatic {
		c.globalVarTypes[globalName] = declaredType
		if declaredType == VTRecord || declaredType == VTObject {
			c.globalRecordTypes[globalName] = udtName // We reuse globalRecordTypes to store Class name for VTObject
		}
	} else if c.isLocal {
		c.localVarTypes[lower] = declaredType
		if declaredType == VTRecord || declaredType == VTObject {
			c.localRecordTypes[lower] = udtName
		}
	} else {
		c.globalVarTypes[lower] = declaredType
		if declaredType == VTRecord || declaredType == VTObject {
			c.globalRecordTypes[lower] = udtName
		}
	}

	// If it's a UDT, we also need to emit an initialization opcode to allocate the record.
	if declaredType == VTRecord {
		udtIdx, ok := c.recordDeclLookup[strings.ToLower(udtName)]
		if ok {
			c.emitExt(ExtOpInitRecord, udtIdx)
			op, idx := c.resolveSetVar(name)
			c.emit(OpSet, int(op), idx) // Use OpSet to assign the record instance
		}
	}
}

func (c *Compiler) parseDimStatement() {
	c.expectKeyword(vbscript.KeywordDim)

	withEvents := false
	if c.checkKeyword(vbscript.KeywordWithEvents) {
		c.move()
		withEvents = true
	}

	for {
		name := c.expectIdentifier()
		c.declareVar(name)

		// Parse optional VB6 As Type clause.
		declaredType, udtName := c.parseAsTypeClause()

		if c.tryParseArrayDeclaration(name) {
			if p, ok := c.next.(*vbscript.PunctuationToken); ok && p.Type == vbscript.PunctEqual {
				panic(c.vbSyntaxError(vbscript.ExpectedEndOfStatement))
			}
		} else {
			if p, ok := c.next.(*vbscript.PunctuationToken); ok && p.Type == vbscript.PunctEqual {
				panic(c.vbSyntaxError(vbscript.ExpectedEndOfStatement))
			}
		}

		// Emit type initialization opcode if As Type was specified.
		c.emitTypedInit(name, declaredType, udtName)

		if withEvents {
			c.emitExt(ExtOpWithEventsRegister, 0xFFFF, c.addConstant(NewString(name)))
		}

		if p, ok := c.next.(*vbscript.PunctuationToken); ok && p.Type == vbscript.PunctComma {
			c.move()
		} else {
			break
		}
	}
}

// parseEraseStatement compiles one VBScript Erase statement for an existing variable.
func (c *Compiler) parseEraseStatement() {
	c.expectKeyword(vbscript.KeywordErase)
	c.parseEraseStatementAfterNameToken()
}

// parseEraseStatementAfterNameToken compiles Erase after the leading token is consumed.
func (c *Compiler) parseEraseStatementAfterNameToken() {
	name := c.expectIdentifier()
	op, idx := c.resolveEraseVar(name)
	c.emit(op, idx)
}

// parseReDimStatement compiles ReDim and ReDim Preserve declarations into runtime array resize helpers.
func (c *Compiler) parseReDimStatement() {
	c.expectKeyword(vbscript.KeywordReDim)
	preserve := false
	if c.checkKeyword(vbscript.KeywordPreserve) {
		c.move()
		preserve = true
	}

	for {
		name := c.expectIdentifier()
		if !(c.isLocal && c.currentClassName != "" && c.hasClassFieldDeclaration(c.currentClassName, name)) {
			c.declareVar(name)
		}

		argCount, ok := c.parseArrayBoundsIntoBuiltinCall(name, preserve)
		if !ok {
			panic(c.vbCompileError(vbscript.ExpectedLParen, "Expected array bounds after ReDim variable name"))
		}
		_ = argCount
		c.emitSetForName(name)

		if p, ok := c.next.(*vbscript.PunctuationToken); ok && p.Type == vbscript.PunctComma {
			c.move()
			continue
		}
		break
	}
}

// tryParseArrayDeclaration compiles Dim array bounds when a declaration includes parentheses.
func (c *Compiler) tryParseArrayDeclaration(name string) bool {
	_, ok := c.parseArrayBoundsIntoBuiltinCall(name, false)
	if !ok {
		return false
	}
	c.emitSetForName(name)
	return true
}

// parseArrayBoundsIntoBuiltinCall emits a helper builtin call for Dim or ReDim array allocation.
func (c *Compiler) parseArrayBoundsIntoBuiltinCall(name string, preserve bool) (int, bool) {
	openParen, ok := c.next.(*vbscript.PunctuationToken)
	if !ok || openParen.Type != vbscript.PunctLParen {
		return 0, false
	}

	if preserve {
		c.emitBuiltinTarget("__AXON_REDIM_PRESERVE_ARRAY_VB6")
		op, idx := c.resolveVar(name)
		c.emit(op, idx)
	} else {
		c.emitBuiltinTarget("__AXON_DIM_ARRAY_VB6")
	}

	c.move()
	argCount := 0
	if closeParen, ok := c.next.(*vbscript.PunctuationToken); !ok || closeParen.Type != vbscript.PunctRParen {
		for {
			c.parseExpression(PrecNone)
			// Handle VB6 A To B syntax
			if c.checkKeyword(vbscript.KeywordTo) {
				c.move()
				c.parseExpression(PrecNone)
				argCount += 2
			} else {
				// Classic VBScript Dim arr(N) uses Option Base as lower bound
				// Push default lower bound AFTER the upper bound, then swap.
				c.emit(OpConstant, c.addConstant(NewInteger(int64(c.optionBase))))
				c.emit(OpSwap)
				argCount += 2
			}
			if comma, ok := c.next.(*vbscript.PunctuationToken); ok && comma.Type == vbscript.PunctComma {
				c.move()
				continue
			}
			break
		}
	}

	closeParen, ok := c.next.(*vbscript.PunctuationToken)
	if !ok || closeParen.Type != vbscript.PunctRParen {
		panic(c.vbCompileError(vbscript.ExpectedRParen, "Expected ')' after array bounds"))
	}
	c.move()

	if preserve {
		c.emit(OpCall, argCount+1)
	} else {
		c.emit(OpCall, argCount)
	}
	return argCount, true
}

// emitSetForName writes the top-of-stack value into the target variable.
func (c *Compiler) emitSetForName(name string) {
	op, idx := c.resolveSetVar(name)
	c.emit(op, idx)
}

// letOpCode maps a Set opcode to its Let counterpart so that plain
// "name = value" assignments remain distinct from explicit object-reference
// Set assignments. "Set name = expr" keeps the raw Set opcodes.
func (c *Compiler) letOpCode(op OpCode) OpCode {
	switch op {
	case OpSetGlobal:
		return OpLetGlobal
	case OpSetLocal:
		return OpLetLocal
	case OpSetClassMember:
		return OpLetClassMember
	default:
		return op
	}
}

// emitBuiltinTarget pushes a builtin function target onto the stack.
func (c *Compiler) emitBuiltinTarget(name string) {
	op, idx := c.resolveVar(name)
	c.emit(op, idx)
}

// compileImplicitClassStatementCall compiles same-class method calls used in statement position.
func (c *Compiler) compileImplicitClassStatementCall(name string, hasParen bool) bool {
	if c == nil || c.currentClassName == "" || (!c.isLocal && !c.dynamicMemberResolution) {
		return false
	}
	trimmedName := strings.TrimSpace(name)
	if trimmedName == "" || strings.EqualFold(trimmedName, c.currentFunctionName) {
		return false
	}
	if globalIdx, exists := c.Globals.Get(trimmedName); exists && globalIdx < c.userGlobalsStart {
		// Keep ASP intrinsics/VBScript builtins and constants (pre-user-global slots)
		// bound as globals inside class methods; do not rewrite them as Me.<member>().
		return false
	}
	if _, exists := c.locals.Get(trimmedName); exists {
		return false
	}
	if _, exists := BuiltinIndex[strings.ToLower(trimmedName)]; exists {
		return false
	}
	if c.hasClassFieldDeclaration(c.currentClassName, trimmedName) {
		return false
	}
	// Allow forward same-class statement calls before later methods are parsed.
	// This matches common VBScript class patterns such as helper calls declared later.

	c.emit(OpMe)
	if hasParen {
		argCount := c.parseParenArgumentList()
		midx := c.addConstant(NewString(trimmedName))
		c.emit(OpCallMember, midx, argCount)
		c.emit(OpPop)
		return true
	}

	argCount := 0
	if !c.isStatementEnd() {
		for {
			if comma, ok := c.next.(*vbscript.PunctuationToken); ok && comma.Type == vbscript.PunctComma {
				emptyIdx := c.addConstant(NewEmpty())
				c.emit(OpConstant, emptyIdx)
			} else if c.isStatementEnd() {
				emptyIdx := c.addConstant(NewEmpty())
				c.emit(OpConstant, emptyIdx)
			} else {
				argStartPos := len(c.bytecode)
				c.parseExpression(PrecNone)
				c.patchArgRefInBytecode(argStartPos)
			}
			argCount++
			if p, ok := c.next.(*vbscript.PunctuationToken); ok && p.Type == vbscript.PunctComma {
				c.move()
				continue
			}
			break
		}
	}

	midx := c.addConstant(NewString(trimmedName))
	c.emit(OpCallMember, midx, argCount)
	c.emit(OpPop)
	return true
}

// parseIfConditionalBlock compiles one If/ElseIf block body and stops before Else, ElseIf, or End.
func (c *Compiler) parseIfConditionalBlock() {
	for !c.matchEof() && !c.checkKeyword(vbscript.KeywordElse) && !c.checkKeyword(vbscript.KeywordElseIf) && !c.checkKeyword(vbscript.KeywordEnd) {
		c.parseStatement()
	}
}

// parseIfElseBlock compiles one Else block body and stops before End.
func (c *Compiler) parseIfElseBlock() {
	for !c.matchEof() && !c.checkKeyword(vbscript.KeywordEnd) {
		c.parseStatement()
	}
}

// parseInlineIfBranchStatements parses one single-line If branch body.
// In VBScript, colon-separated statements after Then/Else/ElseIf belong to that branch.
func (c *Compiler) parseInlineIfBranchStatements() {
	isInlineBoundary := func() bool {
		if c.matchEof() || c.checkKeyword(vbscript.KeywordElseIf) || c.checkKeyword(vbscript.KeywordElse) || c.checkKeyword(vbscript.KeywordEnd) {
			return true
		}
		switch c.next.(type) {
		case *vbscript.LineTerminationToken, *vbscript.CommentToken, *vbscript.ASPCodeEndToken:
			return true
		}
		return false
	}

	for !isInlineBoundary() {
		c.parseStatement()
	}
}

func (c *Compiler) parseIfStatement() {
	c.expectKeyword(vbscript.KeywordIf)
	c.parseExpression(PrecNone)
	c.expectKeyword(vbscript.KeywordThen)
	jumpEndOffsets := make([]int, 0, 2)

	if !c.isStatementEnd() {
		jumpFalseOffset := c.emitJump(OpJumpIfFalse)
		c.parseInlineIfBranchStatements()

		for c.checkKeyword(vbscript.KeywordElseIf) {
			jumpEndOffsets = append(jumpEndOffsets, c.emitJump(OpJump))
			c.patchJump(jumpFalseOffset)

			c.move()
			c.parseExpression(PrecNone)
			c.expectKeyword(vbscript.KeywordThen)

			jumpFalseOffset = c.emitJump(OpJumpIfFalse)
			c.parseInlineIfBranchStatements()
		}

		if c.checkKeyword(vbscript.KeywordElse) {
			c.move()
			jumpEndOffsets = append(jumpEndOffsets, c.emitJump(OpJump))
			c.patchJump(jumpFalseOffset)
			c.parseInlineIfBranchStatements()
		} else {
			c.patchJump(jumpFalseOffset)
		}

		for _, jumpEndOffset := range jumpEndOffsets {
			c.patchJump(jumpEndOffset)
		}

		// Microsoft VBScript compatibility: in single-line If forms, an explicit
		// trailing "End If" is accepted on the same logical line (e.g. "If x Then y=1 : End If").
		// Line terminators are NOT consumed here — consuming them would incorrectly
		// eat the "End" from "End Function", "End Sub", etc. on the following line.
		for {
			switch c.next.(type) {
			case *vbscript.ColonLineTerminationToken, *vbscript.CommentToken:
				c.move()
				continue
			}
			break
		}
		if c.checkKeyword(vbscript.KeywordEnd) {
			c.move()
			if c.checkKeyword(vbscript.KeywordIf) {
				c.move()
			}
		}
		return
	}

	jumpFalseOffset := c.emitJump(OpJumpIfFalse)
	c.parseIfConditionalBlock()

	for c.checkKeyword(vbscript.KeywordElseIf) {
		jumpEndOffsets = append(jumpEndOffsets, c.emitJump(OpJump))
		c.patchJump(jumpFalseOffset)

		c.move()
		c.parseExpression(PrecNone)
		c.expectKeyword(vbscript.KeywordThen)

		jumpFalseOffset = c.emitJump(OpJumpIfFalse)
		if c.isStatementEnd() {
			c.parseIfConditionalBlock()
		} else {
			c.parseStatement()
		}
	}

	if c.checkKeyword(vbscript.KeywordElse) {
		c.move()
		jumpEndOffsets = append(jumpEndOffsets, c.emitJump(OpJump))
		c.patchJump(jumpFalseOffset)
		c.parseIfElseBlock()
	} else {
		c.patchJump(jumpFalseOffset)
	}

	c.expectKeyword(vbscript.KeywordEnd)
	c.expectKeyword(vbscript.KeywordIf)
	for _, jumpEndOffset := range jumpEndOffsets {
		c.patchJump(jumpEndOffset)
	}
}

// parseSelectCaseStatement compiles Select Case ... Case ... End Select blocks.
func (c *Compiler) parseSelectCaseStatement() {
	c.expectKeyword(vbscript.KeywordSelect)
	c.expectKeyword(vbscript.KeywordCase)

	selectValueName := c.newCompilerTempName("select_value")
	c.declareVar(selectValueName)

	c.parseExpression(PrecNone)
	c.emitSetForName(selectValueName)

	jumpEndOffsets := make([]int, 0, 4)
	hasCaseElse := false

	for !c.matchEof() {
		for {
			switch c.next.(type) {
			case *vbscript.LineTerminationToken, *vbscript.ColonLineTerminationToken, *vbscript.CommentToken:
				c.move()
				continue
			}
			break
		}

		if c.checkKeyword(vbscript.KeywordEnd) {
			break
		}

		c.expectKeyword(vbscript.KeywordCase)

		jumpNextCaseOffset := -1
		if c.checkKeyword(vbscript.KeywordElse) {
			hasCaseElse = true
			c.move()
		} else {
			clauseCount := 0
			for {
				opSelectGet, idxSelectGet := c.resolveVar(selectValueName)
				c.emit(opSelectGet, idxSelectGet)

				if c.checkKeyword(vbscript.KeywordIs) {
					// Case Is <comparison-operator> <value>
					c.move() // consume 'Is'

					// Map comparison punctuation to opcode.
					var cmpOp OpCode
					switch pt := c.next.(type) {
					case *vbscript.PunctuationToken:
						switch pt.Type {
						case vbscript.PunctEqual:
							cmpOp = OpEq
						case vbscript.PunctNotEqual:
							cmpOp = OpNeq
						case vbscript.PunctLess:
							cmpOp = OpLt
						case vbscript.PunctGreater:
							cmpOp = OpGt
						case vbscript.PunctLessOrEqual:
							cmpOp = OpLte
						case vbscript.PunctGreaterOrEqual:
							cmpOp = OpGte
						default:
							panic(c.vbCompileError(vbscript.ExpectedEqual, "Expected '='"))
						}
						c.move()
					default:
						panic(c.vbCompileError(vbscript.ExpectedEqual, "Expected '='"))
					}
					c.parseExpression(PrecNone)
					c.emit(cmpOp)
				} else {
					c.parseExpression(PrecNone)

					if c.checkKeyword(vbscript.KeywordTo) {
						// Case low To high
						c.move()
						c.emit(OpGte)

						opSelectGet, idxSelectGet = c.resolveVar(selectValueName)
						c.emit(opSelectGet, idxSelectGet)
						c.parseExpression(PrecNone)
						c.emit(OpLte)
						c.emit(OpAnd)
					} else {
						// Case value
						c.emit(OpEq)
					}
				}

				if clauseCount > 0 {
					c.emit(OpOr)
				}
				clauseCount++

				if comma, ok := c.next.(*vbscript.PunctuationToken); ok && comma.Type == vbscript.PunctComma {
					c.move()
					continue
				}
				break
			}
			jumpNextCaseOffset = c.emitJump(OpJumpIfFalse)
		}

		for !c.matchEof() && !c.checkKeyword(vbscript.KeywordCase) && !c.checkKeyword(vbscript.KeywordEnd) {
			c.parseStatement()
		}

		jumpEndOffsets = append(jumpEndOffsets, c.emitJump(OpJump))
		if jumpNextCaseOffset != -1 {
			c.patchJump(jumpNextCaseOffset)
		}

		if hasCaseElse {
			break
		}
	}

	c.expectKeyword(vbscript.KeywordEnd)
	c.expectKeyword(vbscript.KeywordSelect)

	for _, jumpOffset := range jumpEndOffsets {
		c.patchJump(jumpOffset)
	}
}

// compileASPObjectDeclaration emits initialization code for <object runat="server" ...> declarations.
func (c *Compiler) compileASPObjectDeclaration(objectToken *vbscript.ASPObjectToken) {
	if objectToken == nil || strings.TrimSpace(objectToken.ID) == "" {
		return
	}

	c.ObjectDeclarations = append(c.ObjectDeclarations, objectToken)

	progID := strings.TrimSpace(objectToken.ProgID)
	if progID == "" {
		progID = strings.TrimSpace(objectToken.ClassID)
	}

	emitObjectDeclarationValue := func() {
		if progID == "" {
			emptyIdx := c.addConstant(NewEmpty())
			c.emit(OpConstant, emptyIdx)
			return
		}
		markerIdx := c.addConstant(NewString(staticObjectProgIDPrefix + progID))
		c.emit(OpConstant, markerIdx)
	}

	switch {
	case strings.EqualFold(objectToken.Scope, "application"):
		opApplication, idxApplication := c.resolveVar("Application")
		c.emit(opApplication, idxApplication)

		idIdx := c.addConstant(NewString(objectToken.ID))
		c.emit(OpConstant, idIdx)
		emitObjectDeclarationValue()

		staticObjectsIdx := c.addConstant(NewString("StaticObjects"))
		c.emit(OpCallMember, staticObjectsIdx, 2)
		c.emit(OpPop)

	case strings.EqualFold(objectToken.Scope, "session"):
		opSession, idxSession := c.resolveVar("Session")
		c.emit(opSession, idxSession)

		idIdx := c.addConstant(NewString(objectToken.ID))
		c.emit(OpConstant, idIdx)
		emitObjectDeclarationValue()

		staticObjectsIdx := c.addConstant(NewString("StaticObjects"))
		c.emit(OpCallMember, staticObjectsIdx, 2)
		c.emit(OpPop)
	}
}

func (c *Compiler) parseDoStatement() {
	c.expectKeyword(vbscript.KeywordDo)
	c.pushLoopContext("do")

	loopStart := len(c.bytecode)

	preTestMode := 0 // 0 = none, 1 = while, 2 = until
	if c.matchKeywordOrIdentifier(vbscript.KeywordWhile, "while") {
		preTestMode = 1
		c.move()
	} else if c.matchKeywordOrIdentifier(vbscript.KeywordUntil, "until") {
		preTestMode = 2
		c.move()
	}

	jumpLoopEnd := -1
	if preTestMode != 0 {
		c.parseExpression(PrecNone)
		if preTestMode == 2 {
			c.emit(OpNot)
		}
		jumpLoopEnd = c.emitJump(OpJumpIfFalse)
	}

	for !c.matchEof() && !c.checkKeyword(vbscript.KeywordLoop) {
		c.parseStatement()
	}

	c.expectKeyword(vbscript.KeywordLoop)

	postTestMode := 0 // 0 = none, 1 = while, 2 = until
	if c.matchKeywordOrIdentifier(vbscript.KeywordWhile, "while") {
		postTestMode = 1
		c.move()
	} else if c.matchKeywordOrIdentifier(vbscript.KeywordUntil, "until") {
		postTestMode = 2
		c.move()
	}

	if postTestMode != 0 {
		c.parseExpression(PrecNone)
		if postTestMode == 2 {
			c.emit(OpNot)
		}
		jumpPostTestEnd := c.emitJump(OpJumpIfFalse)
		c.emitLoop(loopStart)
		c.patchJump(jumpPostTestEnd)
	} else {
		c.emitLoop(loopStart)
	}

	if jumpLoopEnd != -1 {
		c.patchJump(jumpLoopEnd)
	}
	c.popLoopContextAndPatch(len(c.bytecode))
}

// parseWhileStatement compiles While...WEnd loops.
func (c *Compiler) parseWhileStatement() {
	c.expectKeyword(vbscript.KeywordWhile)

	loopStart := len(c.bytecode)
	c.parseExpression(PrecNone)
	jumpLoopEnd := c.emitJump(OpJumpIfFalse)

	for !c.matchEof() && !c.matchKeywordOrIdentifier(vbscript.KeywordWEnd, "wend") {
		c.parseStatement()
	}

	if c.matchKeywordOrIdentifier(vbscript.KeywordWEnd, "wend") {
		c.move()
	} else {
		panic(c.vbCompileError(c.keywordMessageCode("Expected keyword WEnd"), "Expected keyword WEnd"))
	}

	c.emitLoop(loopStart)
	c.patchJump(jumpLoopEnd)
}

// parseForStatement compiles For...Next and For Each...Next loops.
func (c *Compiler) parseForStatement() {
	c.expectKeyword(vbscript.KeywordFor)
	if c.checkKeyword(vbscript.KeywordEach) {
		c.parseForEachStatement()
		return
	}
	c.parseForToStatement()
}

// parseForEachStatement compiles For Each loops through internal enumerable helpers.
func (c *Compiler) parseForEachStatement() {
	c.expectKeyword(vbscript.KeywordEach)
	c.pushLoopContext("for")
	loopVarName := c.expectIdentifier()
	c.expectKeyword(vbscript.KeywordIn)

	enumName := c.newCompilerTempName("foreach_enum")
	countName := c.newCompilerTempName("foreach_count")
	indexName := c.newCompilerTempName("foreach_index")

	c.declareVar(enumName)
	c.declareVar(countName)
	c.declareVar(indexName)

	c.emitBuiltinTarget("__AXON_ENUM_VALUES")
	c.parseExpression(PrecNone)
	c.undoTrailingCoerce() // strip OpCoerceToValue so native objects (e.g. RequestCollectionValue) reach __AXON_ENUM_VALUES intact
	c.emit(OpCall, 1)
	c.emitSetForName(enumName)

	c.emitBuiltinTarget("__AXON_ENUM_COUNT")
	opEnumGet, idxEnumGet := c.resolveVar(enumName)
	c.emit(opEnumGet, idxEnumGet)
	c.emit(OpCall, 1)
	c.emitSetForName(countName)

	zeroIdx := c.addConstant(NewInteger(0))
	c.emit(OpConstant, zeroIdx)
	c.emitSetForName(indexName)

	loopStart := len(c.bytecode)
	opIdxGet, idxIdxGet := c.resolveVar(indexName)
	c.emit(opIdxGet, idxIdxGet)
	opCountGet, idxCountGet := c.resolveVar(countName)
	c.emit(opCountGet, idxCountGet)
	c.emit(OpLt)
	jumpLoopEnd := c.emitJump(OpJumpIfFalse)

	c.emitBuiltinTarget("__AXON_ENUM_ITEM")
	opEnumGet, idxEnumGet = c.resolveVar(enumName)
	c.emit(opEnumGet, idxEnumGet)
	opIdxGet, idxIdxGet = c.resolveVar(indexName)
	c.emit(opIdxGet, idxIdxGet)
	c.emit(OpCall, 2)
	c.emitSetForName(loopVarName)

	for !c.matchEof() && !c.checkKeyword(vbscript.KeywordNext) {
		c.parseStatement()
	}

	oneIdx := c.addConstant(NewInteger(1))
	opIdxGet, idxIdxGet = c.resolveVar(indexName)
	c.emit(opIdxGet, idxIdxGet)
	c.emit(OpConstant, oneIdx)
	c.emit(OpAdd)
	c.emitSetForName(indexName)

	c.emitLoop(loopStart)
	c.patchJump(jumpLoopEnd)

	c.expectKeyword(vbscript.KeywordNext)
	if c.isIdentifierLikeToken(c.next) {
		c.move()
	}
	c.popLoopContextAndPatch(len(c.bytecode))
}

// detectUnitStepLiteralFromEmission inspects one freshly emitted step expression and
// reports whether it is a compile-time numeric unit step (+1 or -1).
func (c *Compiler) detectUnitStepLiteralFromEmission(exprStart int) (bool, int64) {
	if c == nil || exprStart < 0 || exprStart >= len(c.bytecode) {
		return false, 0
	}

	emitted := c.bytecode[exprStart:]
	if len(emitted) < 3 || OpCode(emitted[0]) != OpConstant {
		return false, 0
	}

	constIdx := int(binary.BigEndian.Uint16(emitted[1:3]))
	if constIdx < 0 || constIdx >= len(c.constants) {
		return false, 0
	}

	step := int64(0)
	constVal := c.constants[constIdx]
	switch constVal.Type {
	case VTInteger:
		step = constVal.Num
	case VTDouble:
		switch constVal.Flt {
		case 1:
			step = 1
		case -1:
			step = -1
		default:
			return false, 0
		}
	default:
		return false, 0
	}

	if len(emitted) == 4 {
		if OpCode(emitted[3]) != OpNeg {
			return false, 0
		}
		step = -step
	} else if len(emitted) != 3 {
		return false, 0
	}

	if step == 1 || step == -1 {
		return true, step
	}
	return false, 0
}

// emitForLoopStepUpdate emits the loop variable update for one For...Next iteration.
// It uses a dedicated local increment/decrement opcode when the loop variable is local
// and the step is known to be +1 or -1 at compile time.
func (c *Compiler) emitForLoopStepUpdate(loopVarName string, hasUnitStep bool, unitStep int64, stepName string) {
	opLoopGet, idxLoopGet := c.resolveVar(loopVarName)
	if hasUnitStep && opLoopGet == OpGetLocal {
		if unitStep == 1 {
			c.emit(OpIncLocalInt, idxLoopGet)
			return
		}
		if unitStep == -1 {
			c.emit(OpDecLocalInt, idxLoopGet)
			return
		}
	}
	if hasUnitStep && opLoopGet == OpGetGlobal {
		if unitStep == 1 {
			c.emit(OpIncGlobalInt, idxLoopGet)
			return
		}
		if unitStep == -1 {
			c.emit(OpDecGlobalInt, idxLoopGet)
			return
		}
	}

	c.emit(opLoopGet, idxLoopGet)
	opStepGet, idxStepGet := c.resolveVar(stepName)
	c.emit(opStepGet, idxStepGet)
	c.emit(OpAdd)
	c.emitSetForName(loopVarName)
}

// emitForNextFastInt appends the 10-byte OpForNextFastInt super-instruction directly
// into the bytecode slice.  The instruction atomically applies the ±1 step to the
// local counter slot and jumps back to bodyTarget when the counter is still within range.
// stepSign: use +1 for incrementing loops, -1 for decrementing loops.
func (c *Compiler) emitForNextFastInt(varLocalIdx, endLocalIdx int, unitStep int64, bodyTarget int) {
	stepSign := byte(0x01) // +1
	if unitStep < 0 {
		stepSign = 0xFF // -1
	}
	c.bytecode = append(c.bytecode,
		byte(OpForNextFastInt),
		byte(varLocalIdx>>8), byte(varLocalIdx),
		byte(endLocalIdx>>8), byte(endLocalIdx),
		stepSign,
		byte(bodyTarget>>24), byte(bodyTarget>>16), byte(bodyTarget>>8), byte(bodyTarget),
	)
}

// emitForNextFastGlobalInt appends the 10-byte OpForNextFastGlobalInt super-instruction directly
// into the bytecode slice.
func (c *Compiler) emitForNextFastGlobalInt(varGlobalIdx, endGlobalIdx int, unitStep int64, bodyTarget int) {
	stepSign := byte(0x01)
	if unitStep < 0 {
		stepSign = 0xFF
	}
	c.bytecode = append(c.bytecode,
		byte(OpForNextFastGlobalInt),
		byte(varGlobalIdx>>8), byte(varGlobalIdx),
		byte(endGlobalIdx>>8), byte(endGlobalIdx),
		stepSign,
		byte(bodyTarget>>24), byte(bodyTarget>>16), byte(bodyTarget>>8), byte(bodyTarget),
	)
}

// inspectConstantIntEmission returns the compile-time integer value and true if the
// bytecode range [start, end) is exactly one OpConstant instruction that references an
// integer (or whole-number double) constant.  Used by the dead-loop elision pass.
func (c *Compiler) inspectConstantIntEmission(start, end int) (int64, bool) {
	if end-start != 3 {
		return 0, false
	}
	if start < 0 || end > len(c.bytecode) {
		return 0, false
	}
	if OpCode(c.bytecode[start]) != OpConstant {
		return 0, false
	}
	constIdx := int(binary.BigEndian.Uint16(c.bytecode[start+1 : start+3]))
	if constIdx < 0 || constIdx >= len(c.constants) {
		return 0, false
	}
	v := c.constants[constIdx]
	switch v.Type {
	case VTInteger:
		return v.Num, true
	case VTDouble:
		if v.Flt == float64(int64(v.Flt)) {
			return int64(v.Flt), true
		}
	}
	return 0, false
}

// isDeadLoopBody returns true when every opcode in bytecode[start:end] is either
// OpLine (debug location marker) or OpNop (peephole filler).  A body that passes
// this test has no observable side effects and can be elided.
func isDeadLoopBody(bytecode []byte, start, end int) bool {
	for ip := start; ip < end; {
		op := OpCode(bytecode[ip])
		ip++
		switch op {
		case OpLine:
			ip += 4 // [lineH, lineL, colH, colL]
		case OpNop:
			// zero operands
		default:
			return false
		}
	}
	return true
}

// parseForToStatementFastPath compiles the body of a unit-step For...Next loop whose
// counter and limit are both local frame slots.  It emits a simplified single-direction
// pre-loop bounds check and fuses the update+compare+jump tail into one OpForNextFastInt
// super-instruction.  When the body is empty and both bounds are compile-time integer
// constants the entire loop is elided and replaced with a single counter assignment.
func (c *Compiler) parseForToStatementFastPath(
	loopVarName string,
	varLocalIdx, endLocalIdx int,
	unitStep int64,
	initConst int64, initIsConst bool,
	limitConst int64, limitIsConst bool,
) {
	// Record position before the pre-loop check so we can roll back for full elision.
	preLoopStart := len(c.bytecode)

	// Pre-loop check: skip the loop entirely when the range is empty.
	// For +1 step: skip if var > limit  (i.e. jump when NOT var <= limit).
	// For -1 step: skip if var < limit  (i.e. jump when NOT var >= limit).
	c.emit(OpGetLocal, varLocalIdx)
	c.emit(OpGetLocal, endLocalIdx)
	if unitStep == 1 {
		c.emit(OpLte)
	} else {
		c.emit(OpGte)
	}
	jumpExit := c.emitJump(OpJumpIfFalse)

	// Mark the start of the loop body — OpForNextFastInt will jump back here.
	bodyStart := len(c.bytecode)

	// Compile loop body statements.
	for !c.matchEof() && !c.checkKeyword(vbscript.KeywordNext) {
		c.parseStatement()
	}
	bodyEnd := len(c.bytecode)

	// Dead-loop elision: when the body has no observable side effects AND both bounds are
	// compile-time integer constants, replace the entire loop (pre-loop check included) with
	// a single assignment of the post-loop counter value.  This runs in O(1) regardless of
	// the loop range.
	if isDeadLoopBody(c.bytecode, bodyStart, bodyEnd) && initIsConst && limitIsConst {
		// Roll back all emitted bytecode to before the pre-loop check.
		c.bytecode = c.bytecode[:preLoopStart]
		// jumpExit no longer refers to live bytecode; do NOT call patchJump on it.

		// Determine the post-loop counter value.
		var finalVal int64
		loopWillRun := (unitStep == 1 && initConst <= limitConst) ||
			(unitStep == -1 && initConst >= limitConst)
		if loopWillRun {
			// Counter advances one step past the limit on normal completion.
			finalVal = limitConst + unitStep
		} else {
			// Range is empty; counter keeps its initial value.
			finalVal = initConst
		}

		finalIdx := c.addConstant(NewInteger(finalVal))
		c.emit(OpConstant, finalIdx)
		c.emit(OpLetLocal, varLocalIdx)

		c.expectKeyword(vbscript.KeywordNext)
		if c.isIdentifierLikeToken(c.next) {
			c.move()
		}
		c.popLoopContextAndPatch(len(c.bytecode))
		return
	}

	// Normal fast path: fused tail replaces the generic update + direction-check sequence.
	c.emitForNextFastInt(varLocalIdx, endLocalIdx, unitStep, bodyStart)

	// Patch pre-loop exit jump to here.
	c.patchJump(jumpExit)

	c.expectKeyword(vbscript.KeywordNext)
	if c.isIdentifierLikeToken(c.next) {
		c.move()
	}
	c.popLoopContextAndPatch(len(c.bytecode))
}

// parseForToStatement compiles numeric For...Next loops with optional Step expressions.
func (c *Compiler) parseForToStatement() {
	c.pushLoopContext("for")
	loopVarName := c.expectIdentifier()

	if p, ok := c.next.(*vbscript.PunctuationToken); !ok || p.Type != vbscript.PunctEqual {
		panic(c.vbCompileError(vbscript.ExpectedEqual, "Expected '=' in For loop initializer"))
	}
	c.move()

	// Track the bytecode range for the init expression to enable dead-loop elision.
	initExprStart := len(c.bytecode)
	c.parseExpression(PrecNone)
	initExprEnd := len(c.bytecode)
	c.emitSetForName(loopVarName)

	c.expectKeyword(vbscript.KeywordTo)

	endName := c.newCompilerTempName("for_end")
	stepName := c.newCompilerTempName("for_step")
	c.declareVar(endName)
	c.declareVar(stepName)

	// Track the bytecode range for the limit expression to enable dead-loop elision.
	limitExprStart := len(c.bytecode)
	c.parseExpression(PrecNone)
	limitExprEnd := len(c.bytecode)
	c.emitSetForName(endName)
	hasUnitStep := false
	unitStep := int64(0)

	if c.matchKeywordOrIdentifier(vbscript.KeywordStep, "step") {
		c.move()
		stepExprStart := len(c.bytecode)
		c.parseExpression(PrecNone)
		hasUnitStep, unitStep = c.detectUnitStepLiteralFromEmission(stepExprStart)
	} else {
		oneIdx := c.addConstant(NewInteger(1))
		c.emit(OpConstant, oneIdx)
		hasUnitStep, unitStep = true, 1
	}
	c.emitSetForName(stepName)

	// Resolve the loop variable and limit slots before deciding which path to take.
	opLoopGet, idxLoopGet := c.resolveVar(loopVarName)
	opEndGet, idxEndGet := c.resolveVar(endName)

	// ---- FAST PATH ----
	// Conditions: unit step (±1) AND both counter and limit are local frame slots.
	// Emits a simplified single-direction pre-loop check and a fused OpForNextFastInt
	// tail that eliminates the per-iteration direction test and multi-opcode condition.
	if hasUnitStep && opLoopGet == OpGetLocal && opEndGet == OpGetLocal {
		initConst, initIsConst := c.inspectConstantIntEmission(initExprStart, initExprEnd)
		limitConst, limitIsConst := c.inspectConstantIntEmission(limitExprStart, limitExprEnd)
		c.parseForToStatementFastPath(
			loopVarName, idxLoopGet, idxEndGet, unitStep,
			initConst, initIsConst, limitConst, limitIsConst,
		)
		return
	}
	if hasUnitStep && opLoopGet == OpGetGlobal && opEndGet == OpGetGlobal {
		initConst, initIsConst := c.inspectConstantIntEmission(initExprStart, initExprEnd)
		limitConst, limitIsConst := c.inspectConstantIntEmission(limitExprStart, limitExprEnd)
		c.parseForToStatementFastPathGlobal(
			loopVarName, idxLoopGet, idxEndGet, unitStep,
			initConst, initIsConst, limitConst, limitIsConst,
		)
		return
	}

	// ---- SLOW PATH ----
	// Generic direction check: handles non-unit steps, global variables, and float loops.
	loopStart := len(c.bytecode)

	opStepGet, idxStepGet := c.resolveVar(stepName)
	c.emit(opStepGet, idxStepGet)
	zeroIdx := c.addConstant(NewInteger(0))
	c.emit(OpConstant, zeroIdx)
	c.emit(OpGte)
	jumpNegativeCheck := c.emitJump(OpJumpIfFalse)

	c.emit(opLoopGet, idxLoopGet)
	c.emit(opEndGet, idxEndGet)
	c.emit(OpLte)
	jumpLoopEndPositive := c.emitJump(OpJumpIfFalse)
	jumpBody := c.emitJump(OpJump)

	c.patchJump(jumpNegativeCheck)
	opLoopGet2, idxLoopGet2 := c.resolveVar(loopVarName)
	opEndGet2, idxEndGet2 := c.resolveVar(endName)
	c.emit(opLoopGet2, idxLoopGet2)
	c.emit(opEndGet2, idxEndGet2)
	c.emit(OpGte)
	jumpLoopEndNegative := c.emitJump(OpJumpIfFalse)

	c.patchJump(jumpBody)

	for !c.matchEof() && !c.checkKeyword(vbscript.KeywordNext) {
		c.parseStatement()
	}

	c.emitForLoopStepUpdate(loopVarName, hasUnitStep, unitStep, stepName)
	c.emitLoop(loopStart)

	c.patchJump(jumpLoopEndPositive)
	c.patchJump(jumpLoopEndNegative)

	c.expectKeyword(vbscript.KeywordNext)
	if c.isIdentifierLikeToken(c.next) {
		c.move()
	}
	c.popLoopContextAndPatch(len(c.bytecode))
}

// parseSubFunction compiles a Sub or Function declaration and skips body execution at top-level.
func (c *Compiler) parseSubFunction(isFunc bool) {
	if isFunc {
		c.expectKeyword(vbscript.KeywordFunction)
	} else {
		c.expectKeyword(vbscript.KeywordSub)
	}

	name := c.expectIdentifier()
	paramResult := c.parseProcedureParameterNames()
	if isFunc && len(paramResult.names) == 0 {
		c.globalZeroArgFuncs[strings.ToLower(name)] = true
	} else if !isFunc && len(paramResult.names) == 0 {
		c.globalZeroArgSubs[strings.ToLower(name)] = true
	}

	nameIdx, ok := c.Globals.Get(name)
	if !ok {
		nameIdx = c.Globals.Add(name)
	}
	// Sub/Function names must be recognized as declared identifiers so that
	// Option Explicit does not reject forward or backward calls to them.
	if !c.isLocal {
		c.declaredGlobals[strings.ToLower(name)] = true
	}

	placeholder := c.addConstant(NewEmpty())
	c.emit(OpConstant, placeholder)
	c.emit(OpSetGlobal, nameIdx)
	jumpOverBody := c.emitJump(OpJump)

	entryPoint := len(c.bytecode)
	c.constants[placeholder] = NewUserSubEx(entryPoint, len(paramResult.names), c.locals.Count(), isFunc, paramResult.byRefMask, paramResult.optionalMask, paramResult.paramArrayIdx, nil)

	// Store default value constant indices for Optional parameters.
	if len(paramResult.defaults) > 0 {
		defaults := make([]int, len(paramResult.defaults))
		copy(defaults, paramResult.defaults)
		c.funcParamDefaults[entryPoint] = defaults
	}

	c.patchForwardCallSites(name, placeholder)

	prevIsLocal := c.isLocal
	prevLocals := c.locals
	prevDeclared := c.declaredLocals
	prevConstLocals := c.constLocals
	prevStaticLocals := c.staticLocals
	prevFunctionName := c.currentFunctionName
	prevLabelMap := c.labelMap
	prevForwardLabelPatches := c.forwardLabelPatches

	c.isLocal = true
	c.locals = NewSymbolTable()
	c.declaredLocals = make(map[string]bool)
	c.constLocals = make(map[string]bool)
	c.staticLocals = make(map[string]int)
	c.labelMap = make(map[string]int)
	c.forwardLabelPatches = make(map[string][]int)
	c.currentFunctionName = name

	for _, p := range paramResult.names {
		c.locals.Add(p)
		c.declaredLocals[strings.ToLower(p)] = true
	}

	returnIdx := -1
	if isFunc {
		returnIdx = c.locals.Add(name)
		c.declaredLocals[strings.ToLower(name)] = true
	}

	c.hoistProcedureDimDeclarations(keywordFromBool(isFunc))

	for !c.matchEof() {
		if c.checkKeyword(vbscript.KeywordEnd) {
			break
		}
		c.parseStatement()
	}

	c.expectKeyword(vbscript.KeywordEnd)
	if isFunc {
		c.expectKeyword(vbscript.KeywordFunction)
	} else {
		c.expectKeyword(vbscript.KeywordSub)
	}

	c.constants[placeholder] = NewUserSubEx(entryPoint, len(paramResult.names), c.locals.Count(), isFunc, paramResult.byRefMask, paramResult.optionalMask, paramResult.paramArrayIdx, c.locals.names)

	if len(c.forwardLabelPatches) > 0 {
		for label := range c.forwardLabelPatches {
			panic(c.vbCompileError(vbscript.SyntaxError, fmt.Sprintf("Label '%s' not defined in procedure '%s'", label, name)))
		}
	}

	if isFunc {
		c.emit(OpGetLocal, returnIdx)
	} else {
		emptyIdx := c.addConstant(NewEmpty())
		c.emit(OpConstant, emptyIdx)
	}
	c.emit(OpRet)

	c.patchJump(jumpOverBody)

	c.isLocal = prevIsLocal
	c.locals = prevLocals
	c.declaredLocals = prevDeclared
	c.constLocals = prevConstLocals
	c.staticLocals = prevStaticLocals
	c.currentFunctionName = prevFunctionName
	c.labelMap = prevLabelMap
	c.forwardLabelPatches = prevForwardLabelPatches
}

// hoistProcedureDimDeclarations predeclares late Dim names so one procedure keeps
// VBScript whole-procedure local scope even when Dim appears after earlier use.
func (c *Compiler) hoistProcedureDimDeclarations(endKeyword vbscript.Keyword) {
	if c == nil || !c.isLocal || c.next == nil || c.lexer == nil {
		return
	}

	scan := *c.lexer
	tok := c.next
	for tok != nil {
		switch t := tok.(type) {
		case *vbscript.EOFToken:
			return
		case *vbscript.KeywordToken:
			switch t.Keyword {
			case vbscript.KeywordDim:
				c.scanProcedureDimNames(&scan, false)
			case vbscript.KeywordStatic:
				c.scanProcedureDimNames(&scan, true) // Static also reserves names at hoist time
			case vbscript.KeywordEnd:
				nextTok := scan.NextToken()
				if c.tokenMatchesKeywordOrIdentifier(nextTok, endKeyword, strings.ToLower(endKeyword.String())) {
					return
				}
				tok = nextTok
				continue
			}
		}
		tok = scan.NextToken()
	}
}

func keywordFromBool(isFunc bool) vbscript.Keyword {
	if isFunc {
		return vbscript.KeywordFunction
	}
	return vbscript.KeywordSub
}

// scanProcedureDimNames consumes one Dim or Static declaration list from a procedure-body
// scan and predeclares each variable name in the current scope.
func (c *Compiler) scanProcedureDimNames(scan *vbscript.Lexer, isStatic bool) {
	if c == nil || scan == nil {
		return
	}
	for {
		tok := scan.NextToken()
		var name string
		switch t := tok.(type) {
		case *vbscript.IdentifierToken:
			name = t.Name
		case *vbscript.KeywordOrIdentifierToken:
			name = t.Name
		case *vbscript.ExtendedIdentifierToken:
			name = strings.TrimSuffix(strings.TrimPrefix(t.Name, "["), "]")
		default:
			return
		}

		if strings.TrimSpace(name) != "" {
			if isStatic {
				lower := strings.ToLower(name)
				// Hidden global name: __static_[ClassName_]FuncName_varName
				prefix := ""
				if c.currentClassName != "" {
					prefix = c.currentClassName + "_"
				}
				hiddenName := fmt.Sprintf("__static_%s%s_%s", prefix, c.currentFunctionName, name)
				globalIdx := c.Globals.Add(hiddenName)
				c.staticLocals[lower] = globalIdx
				c.declaredGlobals[strings.ToLower(hiddenName)] = true
			} else {
				c.declareVar(name)
			}
		}

		nextTok := scan.NextToken()
		if punct, ok := nextTok.(*vbscript.PunctuationToken); ok && punct.Type == vbscript.PunctLParen {
			depth := 1
			for depth > 0 {
				tok = scan.NextToken()
				if punct, ok := tok.(*vbscript.PunctuationToken); ok {
					switch punct.Type {
					case vbscript.PunctLParen:
						depth++
					case vbscript.PunctRParen:
						depth--
					}
				}
				if _, ok := tok.(*vbscript.EOFToken); ok {
					return
				}
			}
			nextTok = scan.NextToken()
		}

		if punct, ok := nextTok.(*vbscript.PunctuationToken); ok && punct.Type == vbscript.PunctComma {
			continue
		}
		return
	}
}

// newCompilerTempName allocates a deterministic, low-collision temporary variable name.
func (c *Compiler) newCompilerTempName(prefix string) string {
	c.tempCounter++
	return "__axon_" + prefix + "_" + fmt.Sprintf("%d", c.tempCounter)
}

// matchKeywordOrIdentifier checks whether the next token matches one keyword token or keyword-like identifier text.
func (c *Compiler) matchKeywordOrIdentifier(kw vbscript.Keyword, text string) bool {
	if c.checkKeyword(kw) {
		return true
	}
	token := c.next
	if token == nil {
		return false
	}
	switch t := token.(type) {
	case *vbscript.KeywordOrIdentifierToken:
		return strings.EqualFold(t.Name, text)
	case *vbscript.IdentifierToken:
		return strings.EqualFold(t.Name, text)
	default:
		return false
	}
}

func (c *Compiler) matchPunctuation(p vbscript.Punctuation) bool {
	if pt, ok := c.next.(*vbscript.PunctuationToken); ok && pt.Type == p {
		return true
	}
	return false
}

// isIdentifierLikeToken reports whether one token can legally appear where an optional identifier is accepted.
func (c *Compiler) isIdentifierLikeToken(token vbscript.Token) bool {
	switch token.(type) {
	case *vbscript.IdentifierToken, *vbscript.KeywordOrIdentifierToken, *vbscript.ExtendedIdentifierToken:
		return true
	default:
		return false
	}
}

func (c *Compiler) expectKeyword(kw vbscript.Keyword) {
	if k, ok := c.next.(*vbscript.KeywordToken); ok && k.Keyword == kw {
		c.move()
		return
	}
	panic(c.vbCompileError(c.keywordMessageCode(fmt.Sprintf("Expected keyword %v", kw)), fmt.Sprintf("Expected keyword %v", kw)))
}

func (c *Compiler) checkKeyword(kw vbscript.Keyword) bool {
	if k, ok := c.next.(*vbscript.KeywordToken); ok && k.Keyword == kw {
		return true
	}
	if k, ok := c.next.(*vbscript.KeywordOrIdentifierToken); ok && k.Keyword == kw {
		return true
	}
	return false
}

func (c *Compiler) expectIdentifier() string {
	switch t := c.next.(type) {
	case *vbscript.IdentifierToken:
		c.move()
		return t.Name
	case *vbscript.KeywordOrIdentifierToken:
		c.move()
		return t.Name
	case *vbscript.KeywordToken:
		c.move()
		return t.Keyword.String()
	case *vbscript.ExtendedIdentifierToken:
		c.move()
		return strings.TrimSuffix(strings.TrimPrefix(t.Name, "["), "]")
	default:
		panic(c.vbCompileError(vbscript.ExpectedIdentifier, fmt.Sprintf("Expected identifier, got %T", c.next)))
	}
}

// parseResponseWriteFlatChain parses the top-level & concatenation chain that
// appears as the argument to Response.Write and pushes each operand individually
// onto the VM stack.  It stops as soon as the token stream no longer has a '&'
// operator at the concatenation precedence level (i.e. it honours all higher-
// priority operators such as +, *, unary Not, etc.).
//
// The returned count tells OpWriteN how many stack values to consume.
// Only the outmost & chain is flattened; nested parenthesised expressions or

// parseForToStatementFastPathGlobal compiles the body of a unit-step For...Next loop whose
// counter and limit are both global slots.
func (c *Compiler) parseForToStatementFastPathGlobal(
	loopVarName string,
	varGlobalIdx, endGlobalIdx int,
	unitStep int64,
	initConst int64, initIsConst bool,
	limitConst int64, limitIsConst bool,
) {
	preLoopStart := len(c.bytecode)

	c.emit(OpGetGlobal, varGlobalIdx)
	c.emit(OpGetGlobal, endGlobalIdx)
	if unitStep == 1 {
		c.emit(OpLte)
	} else {
		c.emit(OpGte)
	}
	jumpExit := c.emitJump(OpJumpIfFalse)

	bodyStart := len(c.bytecode)

	for !c.matchEof() && !c.checkKeyword(vbscript.KeywordNext) {
		c.parseStatement()
	}
	bodyEnd := len(c.bytecode)

	if isDeadLoopBody(c.bytecode, bodyStart, bodyEnd) && initIsConst && limitIsConst {
		c.bytecode = c.bytecode[:preLoopStart]

		var finalVal int64
		loopWillRun := (unitStep == 1 && initConst <= limitConst) ||
			(unitStep == -1 && initConst >= limitConst)
		if loopWillRun {
			finalVal = limitConst + unitStep
		} else {
			finalVal = initConst
		}

		finalIdx := c.addConstant(NewInteger(finalVal))
		c.emit(OpConstant, finalIdx)
		c.emit(OpLetGlobal, varGlobalIdx)

		c.expectKeyword(vbscript.KeywordNext)
		if c.isIdentifierLikeToken(c.next) {
			c.move()
		}
		c.popLoopContextAndPatch(len(c.bytecode))
		return
	}

	c.emitForNextFastGlobalInt(varGlobalIdx, endGlobalIdx, unitStep, bodyStart)

	c.patchJump(jumpExit)

	c.expectKeyword(vbscript.KeywordNext)
	if c.isIdentifierLikeToken(c.next) {
		c.move()
	}
	c.popLoopContextAndPatch(len(c.bytecode))
}

// sub-expressions with their own & are compiled normally (they still produce a
// single Value via the regular OpConcat path inside parseExpression).
func (c *Compiler) expectStatementEnd() {
	if c.isStatementEnd() {
		// Only move if it's not EOF or CodeEnd (which we want to keep for the main loop)
		switch c.next.(type) {
		case *vbscript.LineTerminationToken, *vbscript.ColonLineTerminationToken, *vbscript.CommentToken:
			c.move()
		}
		return
	}
	panic(c.vbCompileError(vbscript.SyntaxError, "Expected end of statement"))
}

func (c *Compiler) parseTypeDeclaration() {
	if c.isLocal || c.currentClassName != "" {
		panic(c.vbCompileError(vbscript.SyntaxError, "Type declaration is only allowed in global scope"))
	}
	if strings.EqualFold(c.nextIdentifierName(), "type") {
		c.move() // Consume "Type" when it is still the current token.
	}
	typeName := c.expectIdentifier()
	lowerTypeName := strings.ToLower(typeName)

	if _, exists := c.recordDeclLookup[lowerTypeName]; exists {
		panic(c.vbCompileError(vbscript.SyntaxError, fmt.Sprintf("Type '%s' already defined", typeName)))
	}

	decl := CompiledRecordDecl{
		Name:    typeName,
		Members: make([]CompiledRecordMemberDecl, 0),
	}

	c.expectStatementEnd()

	for !c.matchEof() {
		switch c.next.(type) {
		case *vbscript.LineTerminationToken, *vbscript.ColonLineTerminationToken, *vbscript.CommentToken:
			c.move()
			continue
		}
		if c.matchKeywordOrIdentifier(vbscript.KeywordEnd, "end") {
			c.move()
			if strings.EqualFold(c.nextIdentifierName(), "type") {
				c.move()
				break
			}
			panic(c.vbCompileError(vbscript.ExpectedEnd, "Expected 'End Type'"))
		}

		memberName := c.expectIdentifier()
		declaredType, udtName := c.parseAsTypeClause()

		decl.Members = append(decl.Members, CompiledRecordMemberDecl{
			Name:    memberName,
			Type:    declaredType,
			UDTName: udtName,
		})
		c.expectStatementEnd()
	}

	c.recordDeclLookup[lowerTypeName] = len(c.recordDecls)
	c.recordDecls = append(c.recordDecls, decl)
}

func (c *Compiler) parseResponseWriteFlatChain() int {
	// Parse the first operand using PrecConcat so arithmetic binds before the
	// outermost & chain while keeping the chain itself unconsumed here.
	c.parseExpression(PrecConcat)
	count := 1
	for {
		if p, ok := c.next.(*vbscript.PunctuationToken); !ok || p.Type != vbscript.PunctAmp {
			break
		}
		c.move() // consume '&'
		c.parseExpression(PrecConcat)
		count++
	}
	return count
}

// parseClassEventDeclaration parses an Event declaration within a Class.
func (c *Compiler) parseClassEventDeclaration(className string) {
	c.move() // consume "Event"
	eventName := c.expectIdentifier()

	// Events can have parameters in VB6, but for now we'll support empty signatures
	// or just parse and ignore them to match compatibility if needed.
	// Classic ASP usually doesn't have events, so we're adding this for VB6 modernization.
	if p, ok := c.next.(*vbscript.PunctuationToken); ok && p.Type == vbscript.PunctLParen {
		c.move()
		// Parse parameter list but we don't strictly need it for dispatch yet
		for {
			if c.matchEof() {
				break
			}
			if rp, ok2 := c.next.(*vbscript.PunctuationToken); ok2 && rp.Type == vbscript.PunctRParen {
				break
			}
			c.move() // dummy skip for now
			if comma, ok3 := c.next.(*vbscript.PunctuationToken); ok3 && comma.Type == vbscript.PunctComma {
				c.move()
				continue
			}
			break
		}
		if rp, ok := c.next.(*vbscript.PunctuationToken); !ok || rp.Type != vbscript.PunctRParen {
			panic(c.vbCompileError(vbscript.SyntaxError, "Expected ')'"))
		}
		c.move()
	}

	c.addClassEventDeclaration(className, CompiledClassEventDecl{Name: eventName})

	classNameIdx := c.addConstant(NewString(className))
	eventNameIdx := c.addConstant(NewString(eventName))
	c.emitExt(ExtOpRegisterClassEvent, classNameIdx, eventNameIdx)
}

// parseRaiseEventStatement parses a RaiseEvent statement.
func (c *Compiler) parseRaiseEventStatement() {
	c.expectKeyword(vbscript.KeywordRaiseEvent)
	eventName := c.expectIdentifier()

	argCount := 0
	if p, ok := c.next.(*vbscript.PunctuationToken); ok && p.Type == vbscript.PunctLParen {
		c.move()
		if rp, ok2 := c.next.(*vbscript.PunctuationToken); !(ok2 && rp.Type == vbscript.PunctRParen) {
			for {
				c.parseExpression(PrecNone)
				argCount++
				if comma, ok3 := c.next.(*vbscript.PunctuationToken); ok3 && comma.Type == vbscript.PunctComma {
					c.move()
					continue
				}
				break
			}
		}
		if rp, ok := c.next.(*vbscript.PunctuationToken); !ok || rp.Type != vbscript.PunctRParen {
			panic(c.vbCompileError(vbscript.SyntaxError, "Expected ')'"))
		}
		c.move()
	} else if !c.isStatementEnd() {
		// Support RaiseEvent EventName arg1, arg2
		for {
			c.parseExpression(PrecNone)
			argCount++
			if comma, ok3 := c.next.(*vbscript.PunctuationToken); ok3 && comma.Type == vbscript.PunctComma {
				c.move()
				continue
			}
			break
		}
	}

	eventNameIdx := c.addConstant(NewString(eventName))
	c.emitExt(ExtOpRaiseEvent, eventNameIdx, argCount)
}

// parseImplementsStatement parses an Implements statement within a Class.
func (c *Compiler) parseImplementsStatement(className string) {
	c.expectKeyword(vbscript.KeywordImplements)
	interfaceName := c.expectIdentifier()

	c.addClassInterface(className, interfaceName)

	classNameIdx := c.addConstant(NewString(className))
	interfaceNameIdx := c.addConstant(NewString(interfaceName))
	c.emitExt(ExtOpRegisterClassInterface, classNameIdx, interfaceNameIdx)
}

// parseOpenStatement parses a VB6 Open statement: Open path For mode As #num
func (c *Compiler) parseOpenStatement() {
	c.move() // consume Open
	c.parseOpenStatementAfterOpen()
}

func (c *Compiler) parseOpenStatementAfterOpen() {
	c.parseExpression(PrecNone) // path

	if !c.matchKeywordOrIdentifier(vbscript.KeywordFor, "for") {
		panic(c.vbCompileError(vbscript.SyntaxError, "Expected 'For' in Open statement"))
	}
	c.move()

	mode := 0 // 1:Input, 2:Output, 3:Append, 4:Binary, 5:Random
	modeToken := c.nextIdentifierName()
	switch strings.ToLower(modeToken) {
	case "input":
		mode = 1
	case "output":
		mode = 2
	case "append":
		mode = 3
	case "binary":
		mode = 4
	case "random":
		mode = 5
	default:
		panic(c.vbCompileError(vbscript.SyntaxError, "Invalid file mode: "+modeToken))
	}
	c.move()

	// Handle optional Access and Lock? (e.g. Access Read Write)
	for strings.EqualFold(c.nextIdentifierName(), "access") || strings.EqualFold(c.nextIdentifierName(), "lock") {
		c.move()
		// consume next tokens until 'as'
		for !c.matchKeywordOrIdentifier(vbscript.KeywordAs, "as") && !c.matchEof() && !c.isStatementEnd() {
			c.move()
		}
	}

	if !c.matchKeywordOrIdentifier(vbscript.KeywordAs, "as") {
		panic(c.vbCompileError(vbscript.SyntaxError, "Expected 'As' in Open statement"))
	}
	c.move()

	if p, ok := c.next.(*vbscript.PunctuationToken); ok && p.Type == vbscript.PunctHash {
		c.move() // consume #
	}

	modeIdx := c.addConstant(NewInteger(int64(mode)))
	c.emit(OpConstant, modeIdx)

	c.parseExpression(PrecNone) // file number
	c.emitExt(ExtOpFileOpen)
}

// parsePrintStatement parses a VB6 Print # statement: Print #num, expr1, expr2...
func (c *Compiler) parsePrintStatement() {
	c.move() // consume Print
	c.parsePrintStatementAfterPrint()
}

func (c *Compiler) parsePrintStatementAfterPrint() {
	if p, ok := c.next.(*vbscript.PunctuationToken); ok && p.Type == vbscript.PunctHash {
		c.move()
	}
	c.parseExpression(PrecNone) // file number

	if p, ok := c.next.(*vbscript.PunctuationToken); ok && p.Type == vbscript.PunctComma {
		c.move()
	}

	argCount := 0
	if !c.isStatementEnd() {
		for {
			c.parseExpression(PrecNone)
			argCount++
			if p, ok := c.next.(*vbscript.PunctuationToken); ok && p.Type == vbscript.PunctComma {
				c.move()
				continue
			}
			break
		}
	}

	c.emitExt(ExtOpFilePrint, argCount)
}

// parseWriteStatement parses a VB6 Write # statement: Write #num, expr1, expr2...
func (c *Compiler) parseWriteStatement() {
	c.move() // consume Write
	c.parseWriteStatementAfterWrite()
}

func (c *Compiler) parseWriteStatementAfterWrite() {
	if p, ok := c.next.(*vbscript.PunctuationToken); ok && p.Type == vbscript.PunctHash {
		c.move()
	}
	c.parseExpression(PrecNone) // file number

	if p, ok := c.next.(*vbscript.PunctuationToken); ok && p.Type == vbscript.PunctComma {
		c.move()
	}

	argCount := 0
	if !c.isStatementEnd() {
		for {
			c.parseExpression(PrecNone)
			argCount++
			if p, ok := c.next.(*vbscript.PunctuationToken); ok && p.Type == vbscript.PunctComma {
				c.move()
				continue
			}
			break
		}
	}

	c.emitExt(ExtOpFileWrite, argCount)
}

// parseCloseStatement parses a VB6 Close # statement: Close [#num1, #num2...]
func (c *Compiler) parseCloseStatement() {
	c.move() // consume Close
	c.parseCloseStatementAfterClose()
}

func (c *Compiler) parseCloseStatementAfterClose() {
	if c.isStatementEnd() {
		// Close all
		c.emit(OpConstant, c.addConstant(NewInteger(0)))
		c.emitExt(ExtOpFileClose)
		return
	}

	for {
		if p, ok := c.next.(*vbscript.PunctuationToken); ok && p.Type == vbscript.PunctHash {
			c.move()
		}
		c.parseExpression(PrecNone) // num
		c.emitExt(ExtOpFileClose)

		if p, ok := c.next.(*vbscript.PunctuationToken); ok && p.Type == vbscript.PunctComma {
			c.move()
			continue
		}
		break
	}
}

// parseLineInputStatement parses a VB6 Line Input # statement: Line Input #num, var
func (c *Compiler) parseLineInputStatement() {
	// Line and Input already consumed by parseStatement/dispatch
	if p, ok := c.next.(*vbscript.PunctuationToken); ok && p.Type == vbscript.PunctHash {
		c.move()
	}
	c.parseExpression(PrecNone) // num
	c.emitExt(ExtOpFileLineInput)

	if !c.matchPunctuation(vbscript.PunctComma) {
		panic(c.vbCompileError(vbscript.SyntaxError, "Expected ',' after file number in Line Input statement"))
	}
	c.move()

	name := c.expectIdentifier()
	op, idx := c.resolveSetVar(name)
	c.emit(c.letOpCode(op), idx)
}

// parseInputStatement parses a VB6 Input # statement: Input #num, var1, var2...
func (c *Compiler) parseInputStatement() {
	c.move() // consume Input
	c.parseInputStatementAfterInput()
}

func (c *Compiler) parseInputStatementAfterInput() {
	if p, ok := c.next.(*vbscript.PunctuationToken); ok && p.Type == vbscript.PunctHash {
		c.move()
	}
	c.parseExpression(PrecNone) // num

	if !c.matchPunctuation(vbscript.PunctComma) {
		panic(c.vbCompileError(vbscript.SyntaxError, "Expected ',' after file number in Input statement"))
	}
	c.move()

	for {
		name := c.expectIdentifier()
		c.emit(OpJSDup)               // duplicate file num
		c.emitExt(ExtOpFileLineInput) // Placeholder
		op, idx := c.resolveSetVar(name)
		c.emit(c.letOpCode(op), idx)

		if p, ok := c.next.(*vbscript.PunctuationToken); ok && p.Type == vbscript.PunctComma {
			c.move()
			continue
		}
		break
	}
	c.emit(OpPop) // pop file num
}

// parsePutStatement parses a VB6 Put # statement: Put #num, [pos], value
func (c *Compiler) parsePutStatement() {
	c.move() // consume Put
	c.parsePutStatementAfterPut()
}

func (c *Compiler) parsePutStatementAfterPut() {
	if p, ok := c.next.(*vbscript.PunctuationToken); ok && p.Type == vbscript.PunctHash {
		c.move()
	}
	c.parseExpression(PrecNone) // num

	if !c.matchPunctuation(vbscript.PunctComma) {
		panic(c.vbCompileError(vbscript.SyntaxError, "Expected ',' after file number in Put statement"))
	}
	c.move()

	if p, ok := c.next.(*vbscript.PunctuationToken); ok && p.Type == vbscript.PunctComma {
		// optional pos skipped, push empty
		c.emit(OpConstant, c.addConstant(NewEmpty()))
		c.move()
	} else {
		c.parseExpression(PrecNone) // pos
		if !c.matchPunctuation(vbscript.PunctComma) {
			panic(c.vbCompileError(vbscript.SyntaxError, "Expected ',' after position in Put statement"))
		}
		c.move()
	}

	c.parseExpression(PrecNone) // value
	c.emitExt(ExtOpFilePut)
}

// parseGetStatement parses a VB6 Get # statement: Get #num, [pos], var
func (c *Compiler) parseGetStatement() {
	c.move() // consume Get
	c.parseGetStatementAfterGet()
}

func (c *Compiler) parseGetStatementAfterGet() {
	if p, ok := c.next.(*vbscript.PunctuationToken); ok && p.Type == vbscript.PunctHash {
		c.move()
	}
	c.parseExpression(PrecNone) // num

	if !c.matchPunctuation(vbscript.PunctComma) {
		panic(c.vbCompileError(vbscript.SyntaxError, "Expected ',' after file number in Get statement"))
	}
	c.move()

	if p, ok := c.next.(*vbscript.PunctuationToken); ok && p.Type == vbscript.PunctComma {
		// optional pos skipped, push empty
		c.emit(OpConstant, c.addConstant(NewEmpty()))
		c.move()
	} else {
		c.parseExpression(PrecNone) // pos
		if !c.matchPunctuation(vbscript.PunctComma) {
			panic(c.vbCompileError(vbscript.SyntaxError, "Expected ',' after position in Get statement"))
		}
		c.move()
	}

	c.emitExt(ExtOpFileGet)

	name := c.expectIdentifier()
	op, idx := c.resolveSetVar(name)
	c.emit(c.letOpCode(op), idx)
}
