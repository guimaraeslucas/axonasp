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
		default:
			c.parseExpression(PrecNone)
			c.emit(OpPop)
		}
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

		if strings.EqualFold(name, "Option") {
			c.parseOptionStatementAfterOptionKeyword()
			return
		}

		if strings.EqualFold(name, "Erase") {
			c.parseEraseStatementAfterNameToken()
			return
		}

		if p, ok := c.next.(*vbscript.PunctuationToken); ok && p.Type == vbscript.PunctEqual {
			c.move() // Consume '='
			c.parseExpression(PrecNone)
			op, idx := c.resolveSetVar(name)
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
			op, idx := c.resolveVar(name)
			loadPos := c.emit(op, idx)
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
				if _, ok := c.next.(*vbscript.PunctuationToken); ok {
					// In statement call contexts, chain intermediate zero-arg member gets so
					// patterns like aspl.json.dump(x) call Dump on json() result.
					if lp, ok := c.next.(*vbscript.PunctuationToken); ok && lp.Type == vbscript.PunctLParen {
						for i := 0; i < len(memberChain)-1; i++ {
							intermediateIdx := c.addConstant(NewString(memberChain[i]))
							c.emit(OpConstant, intermediateIdx)
							c.emit(OpMemberGet)
						}
						callMemberName = memberChain[len(memberChain)-1]
					}
				}
			}

			// Property assignment: obj.Prop = value
			if peq, ok := c.next.(*vbscript.PunctuationToken); ok && peq.Type == vbscript.PunctEqual {
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
				} else {
					c.emit(OpCallMember, midx, argCount)
					c.emit(OpPop)
				}
			} else {
				// Member sub call without parentheses: obj.Method arg1, arg2
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
		if isPublic && c.isIdentifierLikeToken(c.next) {
			if isDefaultMember {
				panic(c.vbCompileError(vbscript.ExpectedSub, "Default member must be a Sub, Function, or Property"))
			}
			c.parseClassFieldDeclaration(className, true)
			continue
		}
		if !isPublic && c.isIdentifierLikeToken(c.next) {
			if isDefaultMember {
				panic(c.vbCompileError(vbscript.ExpectedSub, "Default member must be a Sub, Function, or Property"))
			}
			c.parseClassFieldDeclaration(className, false)
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

	for {
		fieldName := c.expectIdentifier()
		c.addClassFieldDeclaration(className, CompiledClassFieldDecl{Name: fieldName, IsPublic: isPublic})

		classNameIdx := c.addConstant(NewString(className))
		fieldNameIdx := c.addConstant(NewString(fieldName))
		isPublicOperand := 0
		if isPublic {
			isPublicOperand = 1
		}
		c.emit(OpRegisterClassField, classNameIdx, fieldNameIdx, isPublicOperand)

		// If the field has array dimensions (e.g., Private arr(5) or Private arr(3,5)),
		// parse each upper-bound expression and emit OpInitClassArrayField so the VM can
		// allocate a pre-sized array for every new class instance.
		if p, ok := c.next.(*vbscript.PunctuationToken); ok && p.Type == vbscript.PunctLParen {
			c.move() // consume '('
			dimCount := 0
			// Empty parens Private arr() means dynamic (undimensioned); emit no init opcode.
			if rp, ok2 := c.next.(*vbscript.PunctuationToken); !(ok2 && rp.Type == vbscript.PunctRParen) {
				for {
					c.parseExpression(PrecNone) // pushes upper-bound value onto the stack
					dimCount++
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
			if dimCount > 0 {
				c.emit(OpInitClassArrayField, classNameIdx, fieldNameIdx, dimCount)
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

// parseProcedureParameterNames parses Sub/Function formal parameter names and modifiers.
// Returns the parameter name list and a byRefMask where bit i = 1 means param i is ByRef.
// VBScript default is ByRef when no modifier is specified.
func (c *Compiler) parseProcedureParameterNames() ([]string, uint64) {
	var params []string
	var byRefMask uint64
	if p, ok := c.next.(*vbscript.PunctuationToken); !ok || p.Type != vbscript.PunctLParen {
		return params, byRefMask
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
		if c.checkKeyword(vbscript.KeywordByRef) {
			c.move()
			isByRef = true
		} else if c.checkKeyword(vbscript.KeywordByVal) {
			c.move()
			isByRef = false
		}
		if c.matchKeywordOrIdentifier(vbscript.KeywordOptional, "optional") {
			c.move()
		}
		paramName := c.expectIdentifier()
		params = append(params, paramName)
		if isByRef && paramIdx < 64 {
			byRefMask |= 1 << uint(paramIdx)
		}
		paramIdx++
		if p, ok := c.next.(*vbscript.PunctuationToken); ok && p.Type == vbscript.PunctComma {
			c.move()
		}
	}
	if p, ok := c.next.(*vbscript.PunctuationToken); ok && p.Type == vbscript.PunctRParen {
		c.move()
	}

	return params, byRefMask
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
	params, byRefMask := c.parseProcedureParameterNames()
	if strings.EqualFold(methodName, "Class_Initialize") || strings.EqualFold(methodName, "Class_Terminate") {
		if len(params) != 0 {
			panic(c.vbCompileError(vbscript.ClassInitializeOrTerminateDoNotHaveArguments, "Class_Initialize/Class_Terminate must not declare arguments"))
		}
	}

	placeholder := c.addConstant(NewEmpty())
	jumpOverBody := c.emitJump(OpJump)

	entryPoint := len(c.bytecode)
	c.constants[placeholder] = NewUserSub(entryPoint, len(params), c.locals.Count(), isFunc, byRefMask, nil)

	prevIsLocal := c.isLocal
	prevLocals := c.locals
	prevDeclared := c.declaredLocals
	prevConstLocals := c.constLocals
	prevClassName := c.currentClassName
	prevFunctionName := c.currentFunctionName

	c.isLocal = true
	c.currentClassName = className
	c.locals = NewSymbolTable()
	c.declaredLocals = make(map[string]bool)
	c.constLocals = make(map[string]bool)
	if isFunc {
		c.currentFunctionName = methodName
	} else {
		c.currentFunctionName = ""
	}

	for _, p := range params {
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

	c.constants[placeholder] = NewUserSub(entryPoint, len(params), c.locals.Count(), isFunc, byRefMask, c.locals.names)

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

	c.isLocal = prevIsLocal
	c.currentClassName = prevClassName
	c.locals = prevLocals
	c.declaredLocals = prevDeclared
	c.constLocals = prevConstLocals
	c.currentFunctionName = prevFunctionName
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
	params, byRefMaskProp := c.parseProcedureParameterNames()

	if (accessorKind == classPropertyAccessorLet || accessorKind == classPropertyAccessorSet) && len(params) == 0 {
		panic(c.vbCompileError(vbscript.PropertySetOrLetMustHaveArguments, "Property Let/Set requires one value parameter"))
	}

	placeholder := c.addConstant(NewEmpty())
	jumpOverBody := c.emitJump(OpJump)

	entryPoint := len(c.bytecode)
	c.constants[placeholder] = NewUserSub(entryPoint, len(params), c.locals.Count(), isFunction, byRefMaskProp, nil)

	prevIsLocal := c.isLocal
	prevLocals := c.locals
	prevDeclared := c.declaredLocals
	prevConstLocals := c.constLocals
	prevClassName := c.currentClassName
	prevFunctionName := c.currentFunctionName

	c.isLocal = true
	c.currentClassName = className
	c.locals = NewSymbolTable()
	c.declaredLocals = make(map[string]bool)
	c.constLocals = make(map[string]bool)
	if isFunction {
		c.currentFunctionName = propertyName
	} else {
		c.currentFunctionName = ""
	}

	for _, p := range params {
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

	c.constants[placeholder] = NewUserSub(entryPoint, len(params), c.locals.Count(), isFunction, byRefMaskProp, c.locals.names)

	if isFunction {
		c.emit(OpGetLocal, returnIdx)
	} else {
		emptyIdx := c.addConstant(NewEmpty())
		c.emit(OpConstant, emptyIdx)
	}
	c.emit(OpRet)

	c.patchJump(jumpOverBody)

	c.registerClassPropertyDeclaration(className, propertyName, isPublic, accessorKind, len(params), placeholder)

	classNameIdx := c.addConstant(NewString(className))
	propertyNameIdx := c.addConstant(NewString(propertyName))
	isPublicOperand := 0
	if isPublic {
		isPublicOperand = 1
	}
	switch accessorKind {
	case classPropertyAccessorGet:
		c.emit(OpRegisterClassPropertyGet, classNameIdx, propertyNameIdx, placeholder, len(params), isPublicOperand)
	case classPropertyAccessorLet:
		c.emit(OpRegisterClassPropertyLet, classNameIdx, propertyNameIdx, placeholder, len(params), isPublicOperand)
	case classPropertyAccessorSet:
		c.emit(OpRegisterClassPropertySet, classNameIdx, propertyNameIdx, placeholder, len(params), isPublicOperand)
	}

	if isDefaultMember {
		defaultNameIdx := c.addConstant(NewString("__default__"))
		switch accessorKind {
		case classPropertyAccessorGet:
			c.emit(OpRegisterClassPropertyGet, classNameIdx, defaultNameIdx, placeholder, len(params), isPublicOperand)
		case classPropertyAccessorLet:
			c.emit(OpRegisterClassPropertyLet, classNameIdx, defaultNameIdx, placeholder, len(params), isPublicOperand)
		case classPropertyAccessorSet:
			c.emit(OpRegisterClassPropertySet, classNameIdx, defaultNameIdx, placeholder, len(params), isPublicOperand)
		}
	}

	c.isLocal = prevIsLocal
	c.currentClassName = prevClassName
	c.locals = prevLocals
	c.declaredLocals = prevDeclared
	c.constLocals = prevConstLocals
	c.currentFunctionName = prevFunctionName
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

// parseScopedVariableDeclaration compiles page-scope Public/Private variable declarations.
// Classic ASP accepts module-level declarations like `Private counter` as aliases for
// page-level variable declarations; they should not be rejected as missing procedures.
func (c *Compiler) parseScopedVariableDeclaration() {
	for {
		name := c.expectIdentifier()
		c.declareVar(name)

		if c.tryParseArrayDeclaration(name) {
			if p, ok := c.next.(*vbscript.PunctuationToken); ok && p.Type == vbscript.PunctEqual {
				panic(c.vbSyntaxError(vbscript.ExpectedEndOfStatement))
			}
		} else {
			if p, ok := c.next.(*vbscript.PunctuationToken); ok && p.Type == vbscript.PunctEqual {
				panic(c.vbSyntaxError(vbscript.ExpectedEndOfStatement))
			}
		}

		if p, ok := c.next.(*vbscript.PunctuationToken); ok && p.Type == vbscript.PunctComma {
			c.move()
			continue
		}
		break
	}
}

func (c *Compiler) parseDimStatement() {
	c.expectKeyword(vbscript.KeywordDim)
	for {
		name := c.expectIdentifier()
		c.declareVar(name)

		if c.tryParseArrayDeclaration(name) {
			if p, ok := c.next.(*vbscript.PunctuationToken); ok && p.Type == vbscript.PunctEqual {
				panic(c.vbSyntaxError(vbscript.ExpectedEndOfStatement))
			}
		} else {
			if p, ok := c.next.(*vbscript.PunctuationToken); ok && p.Type == vbscript.PunctEqual {
				panic(c.vbSyntaxError(vbscript.ExpectedEndOfStatement))
			}
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
		c.emitBuiltinTarget("__AXON_REDIM_PRESERVE_ARRAY")
		op, idx := c.resolveVar(name)
		c.emit(op, idx)
	} else {
		c.emitBuiltinTarget("__AXON_DIM_ARRAY")
	}

	c.move()
	argCount := 0
	if closeParen, ok := c.next.(*vbscript.PunctuationToken); !ok || closeParen.Type != vbscript.PunctRParen {
		for {
			c.parseExpression(PrecNone)
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

func (c *Compiler) parseIfStatement() {
	c.expectKeyword(vbscript.KeywordIf)
	c.parseExpression(PrecNone)
	c.expectKeyword(vbscript.KeywordThen)
	jumpEndOffsets := make([]int, 0, 2)

	if !c.isStatementEnd() {
		jumpFalseOffset := c.emitJump(OpJumpIfFalse)
		c.parseStatement()

		for c.checkKeyword(vbscript.KeywordElseIf) {
			jumpEndOffsets = append(jumpEndOffsets, c.emitJump(OpJump))
			c.patchJump(jumpFalseOffset)

			c.move()
			c.parseExpression(PrecNone)
			c.expectKeyword(vbscript.KeywordThen)

			jumpFalseOffset = c.emitJump(OpJumpIfFalse)
			c.parseStatement()
		}

		if c.checkKeyword(vbscript.KeywordElse) {
			c.move()
			jumpEndOffsets = append(jumpEndOffsets, c.emitJump(OpJump))
			c.patchJump(jumpFalseOffset)
			c.parseStatement()
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

// parseForToStatement compiles numeric For...Next loops with optional Step expressions.
func (c *Compiler) parseForToStatement() {
	c.pushLoopContext("for")
	loopVarName := c.expectIdentifier()

	if p, ok := c.next.(*vbscript.PunctuationToken); !ok || p.Type != vbscript.PunctEqual {
		panic(c.vbCompileError(vbscript.ExpectedEqual, "Expected '=' in For loop initializer"))
	}
	c.move()

	c.parseExpression(PrecNone)
	c.emitSetForName(loopVarName)

	c.expectKeyword(vbscript.KeywordTo)

	endName := c.newCompilerTempName("for_end")
	stepName := c.newCompilerTempName("for_step")
	c.declareVar(endName)
	c.declareVar(stepName)

	c.parseExpression(PrecNone)
	c.emitSetForName(endName)

	if c.matchKeywordOrIdentifier(vbscript.KeywordStep, "step") {
		c.move()
		c.parseExpression(PrecNone)
	} else {
		oneIdx := c.addConstant(NewInteger(1))
		c.emit(OpConstant, oneIdx)
	}
	c.emitSetForName(stepName)

	loopStart := len(c.bytecode)

	opStepGet, idxStepGet := c.resolveVar(stepName)
	c.emit(opStepGet, idxStepGet)
	zeroIdx := c.addConstant(NewInteger(0))
	c.emit(OpConstant, zeroIdx)
	c.emit(OpGte)
	jumpNegativeCheck := c.emitJump(OpJumpIfFalse)

	opLoopGet, idxLoopGet := c.resolveVar(loopVarName)
	c.emit(opLoopGet, idxLoopGet)
	opEndGet, idxEndGet := c.resolveVar(endName)
	c.emit(opEndGet, idxEndGet)
	c.emit(OpLte)
	jumpLoopEndPositive := c.emitJump(OpJumpIfFalse)
	jumpBody := c.emitJump(OpJump)

	c.patchJump(jumpNegativeCheck)
	opLoopGet, idxLoopGet = c.resolveVar(loopVarName)
	c.emit(opLoopGet, idxLoopGet)
	opEndGet, idxEndGet = c.resolveVar(endName)
	c.emit(opEndGet, idxEndGet)
	c.emit(OpGte)
	jumpLoopEndNegative := c.emitJump(OpJumpIfFalse)

	c.patchJump(jumpBody)

	for !c.matchEof() && !c.checkKeyword(vbscript.KeywordNext) {
		c.parseStatement()
	}

	opLoopGet, idxLoopGet = c.resolveVar(loopVarName)
	c.emit(opLoopGet, idxLoopGet)
	opStepGet, idxStepGet = c.resolveVar(stepName)
	c.emit(opStepGet, idxStepGet)
	c.emit(OpAdd)
	c.emitSetForName(loopVarName)
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
	params, byRefMaskSub := c.parseProcedureParameterNames()
	if isFunc && len(params) == 0 {
		c.globalZeroArgFuncs[strings.ToLower(name)] = true
	}

	nameIdx, ok := c.Globals.Get(name)
	if !ok {
		nameIdx = c.Globals.Add(name)
	}

	placeholder := c.addConstant(NewEmpty())
	c.emit(OpConstant, placeholder)
	c.emit(OpSetGlobal, nameIdx)
	jumpOverBody := c.emitJump(OpJump)

	entryPoint := len(c.bytecode)
	c.constants[placeholder] = NewUserSub(entryPoint, len(params), c.locals.Count(), isFunc, byRefMaskSub, nil)
	c.patchForwardCallSites(name, placeholder)

	prevIsLocal := c.isLocal
	prevLocals := c.locals
	prevDeclared := c.declaredLocals
	prevConstLocals := c.constLocals
	prevFunctionName := c.currentFunctionName

	c.isLocal = true
	c.locals = NewSymbolTable()
	c.declaredLocals = make(map[string]bool)
	c.constLocals = make(map[string]bool)
	if isFunc {
		c.currentFunctionName = name
	} else {
		c.currentFunctionName = ""
	}

	for _, p := range params {
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

	c.constants[placeholder] = NewUserSub(entryPoint, len(params), c.locals.Count(), isFunc, byRefMaskSub, c.locals.names)

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
	c.currentFunctionName = prevFunctionName
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
				c.scanProcedureDimNames(&scan)
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

// scanProcedureDimNames consumes one Dim declaration list from a procedure-body
// scan and predeclares each variable name in the current local symbol table.
func (c *Compiler) scanProcedureDimNames(scan *vbscript.Lexer) {
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
			c.declareVar(name)
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
