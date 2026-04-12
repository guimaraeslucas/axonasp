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

type Precedence int

const (
	PrecNone       Precedence = iota
	PrecAssignment            // =
	PrecEqv                   // Eqv
	PrecImp                   // Imp
	PrecXor                   // Xor
	PrecOr                    // Or
	PrecAnd                   // And
	PrecNot                   // Not
	PrecEquality              // = <> < > <= >= Is
	PrecConcat                // &
	PrecTerm                  // + -
	PrecFactor                // * / \ Mod
	PrecUnary                 // + -
	PrecExp                   // ^
	PrecCall                  // . ()
)

// emitIdentifierValue emits an identifier reference as a value expression.
// In class scope, a bare method/property getter name is compiled as implicit Me.<member>()
// when it resolves to a zero-argument callable member.
func (c *Compiler) emitIdentifierValue(name string) {
	trimmedName := strings.TrimSpace(name)
	if c.isLocal && c.currentFunctionName != "" && strings.EqualFold(trimmedName, c.currentFunctionName) {
		if p, ok := c.next.(*vbscript.PunctuationToken); ok && p.Type == vbscript.PunctLParen {
			if idx, exists := c.Globals.Get(trimmedName); exists {
				pos := c.emit(OpGetGlobal, idx)
				c.markLastCallTarget(trimmedName, OpGetGlobal, pos)
				return
			}
		}
	}
	if c.currentClassName != "" && (c.isLocal || c.dynamicMemberResolution) {
		if _, exists := c.locals.Get(trimmedName); !exists {
			if p, ok := c.next.(*vbscript.PunctuationToken); ok && p.Type == vbscript.PunctLParen {
				// Parenthesized member calls are handled by compileImplicitClassMemberCall.
			} else if c.hasClassMethodDeclaration(c.currentClassName, trimmedName) || c.hasClassZeroArgPropertyGetDeclaration(c.currentClassName, trimmedName) {
				c.clearLastCallTarget()
				c.emit(OpMe)
				midx := c.addConstant(NewString(trimmedName))
				c.emit(OpCallMember, midx, 0)
				return
			}
		}
	}

	if p, ok := c.next.(*vbscript.PunctuationToken); !(ok && p.Type == vbscript.PunctLParen) {
		if c.emitStaticObjectIdentifierFallback(trimmedName) {
			return
		}
	}

	op, idx := c.resolveVar(trimmedName)
	pos := c.emit(op, idx)
	c.markLastCallTarget(trimmedName, op, pos)
	if op == OpGetGlobal {
		lower := strings.ToLower(strings.TrimSpace(trimmedName))
		if !c.constGlobals[lower] {
			c.registerForwardConstPatch(trimmedName, pos)
		}
	}

	if op == OpGetGlobal || op == OpGetLocal || op == OpGetClassMember {
		c.emitTrailingCoerceIfValueContext()
	}
}

// emitTrailingCoerceIfValueContext emits OpCoerceToValue only when the current
// expression position is a value context and not part of member/call chaining.
func (c *Compiler) emitTrailingCoerceIfValueContext() {
	var suppressCoerce bool
	if p, ok := c.next.(*vbscript.PunctuationToken); ok {
		suppressCoerce = p.Type == vbscript.PunctDot || p.Type == vbscript.PunctLParen
	}
	if !suppressCoerce {
		if kw, ok := c.next.(*vbscript.KeywordToken); ok {
			suppressCoerce = kw.Keyword == vbscript.KeywordIs
		}
	}
	if !suppressCoerce {
		c.emit(OpCoerceToValue)
	}
}

// undoTrailingCoerce removes the last emitted OpCoerceToValue byte from the bytecode
// when the compiler needs a raw object reference (member access, Is operator, Set).
func (c *Compiler) undoTrailingCoerce() {
	if c.lastCoercePos == len(c.bytecode)-1 && c.lastCoercePos >= 0 {
		c.bytecode = c.bytecode[:len(c.bytecode)-1]
		c.lastCoercePos = -1
	}
}

// tryInlineEvalCall compiles Eval("...") directly in the current scope so the
// evaluated expression resolves the same globals/locals as Microsoft VBScript.
// It only applies when the Eval argument is a single string literal constant.
func (c *Compiler) tryInlineEvalCall(callTargetName string, callTargetPos int, argExprStart int, argCount int) bool {
	if !strings.EqualFold(callTargetName, "Eval") {
		return false
	}
	if argCount != 1 {
		return false
	}
	if callTargetPos < 0 || callTargetPos > len(c.bytecode) {
		return false
	}
	if argExprStart < 0 || argExprStart+3 != len(c.bytecode) {
		return false
	}
	if OpCode(c.bytecode[argExprStart]) != OpConstant {
		return false
	}
	constIdx := int(binary.BigEndian.Uint16(c.bytecode[argExprStart+1:]))
	if constIdx < 0 || constIdx >= len(c.constants) {
		return false
	}
	if c.constants[constIdx].Type != VTString {
		return false
	}

	expr := c.constants[constIdx].Str
	originalBytecode := append([]byte(nil), c.bytecode...)
	originalName := c.lastCallTargetName
	originalPos := c.lastCallTargetPos
	originalIsGlobal := c.lastCallIsGlobal

	// Remove the emitted Eval call target and argument bytes.
	c.bytecode = c.bytecode[:callTargetPos]
	c.clearLastCallTarget()

	if c.compileInlineEvalExpression(expr) {
		return true
	}

	// Restore if inlining fails for any reason, then fall back to runtime Eval.
	c.bytecode = originalBytecode
	c.lastCallTargetName = originalName
	c.lastCallTargetPos = originalPos
	c.lastCallIsGlobal = originalIsGlobal
	return false
}

// compileInlineEvalExpression parses one expression text using the same compiler
// context (globals, locals, options) and emits its bytecode into the current stream.
func (c *Compiler) compileInlineEvalExpression(expr string) (ok bool) {
	if c == nil {
		return false
	}

	originalLexer := c.lexer
	originalNext := c.next
	originalLexerMode := c.lexerMode
	originalSourceCode := c.sourceCode
	originalDebugLine := c.lastDebugLine
	originalDebugColumn := c.lastDebugColumn

	defer func() {
		c.lexer = originalLexer
		c.next = originalNext
		c.lexerMode = originalLexerMode
		c.sourceCode = originalSourceCode
		c.lastDebugLine = originalDebugLine
		c.lastDebugColumn = originalDebugColumn
		if r := recover(); r != nil {
			ok = false
		}
	}()

	inlineLexer := vbscript.NewLexer(expr)
	inlineLexer.Mode = vbscript.ModeVBScript
	inlineLexer.InASPBlock = true

	c.lexer = inlineLexer
	c.lexerMode = vbscript.ModeVBScript
	c.sourceCode = expr
	c.next = nil
	c.lastDebugLine = -1
	c.lastDebugColumn = -1
	c.move()

	c.parseExpression(PrecNone)
	if !c.matchEof() {
		return false
	}

	return true
}

// emitStaticObjectIdentifierFallback compiles undeclared ASP identifier reads as
// Session.StaticObjects(name) with Application.StaticObjects(name) fallback.
func (c *Compiler) emitStaticObjectIdentifierFallback(name string) bool {
	if c == nil || c.isLocal || c.lexerMode != vbscript.ModeASP {
		return false
	}
	if c.optionExplicit {
		return false
	}
	if strings.TrimSpace(name) == "" {
		return false
	}
	if _, exists := c.Globals.Get(name); exists {
		return false
	}

	nameGetOp, nameIdx := c.resolveVar(name)
	nameSetOp, nameSetIdx := c.resolveSetVar(name)
	if nameGetOp != OpGetGlobal || nameSetOp != OpSetGlobal {
		return false
	}

	nameConstIdx := c.addConstant(NewString(name))
	staticObjectsConstIdx := c.addConstant(NewString("StaticObjects"))
	itemConstIdx := c.addConstant(NewString("Item"))
	tmpName := c.newCompilerTempName("staticobj")
	c.declareVar(tmpName)
	useCurrentJump := -1
	if nameGetOp == OpGetGlobal {
		c.emitBuiltinTarget("IsEmpty")
		c.emit(nameGetOp, nameIdx)
		c.emit(OpCall, 1)
		useCurrentJump = c.emitJump(OpJumpIfFalse)
	}

	// tmp = Session.StaticObjects.Item(name)
	opSession, idxSession := c.resolveVar("Session")
	c.emit(opSession, idxSession)
	c.emit(OpCallMember, staticObjectsConstIdx, 0)
	c.emit(OpConstant, nameConstIdx)
	c.emit(OpCallMember, itemConstIdx, 1)
	c.emitSetForName(tmpName)

	// If IsObject(tmp) Then result = tmp Else result = Application.StaticObjects(name)
	c.emitBuiltinTarget("IsObject")
	opTmpGet, idxTmpGet := c.resolveVar(tmpName)
	c.emit(opTmpGet, idxTmpGet)
	c.emit(OpCall, 1)
	jumpUseApplication := c.emitJump(OpJumpIfFalse)

	opTmpGet, idxTmpGet = c.resolveVar(tmpName)
	c.emit(opTmpGet, idxTmpGet)
	c.emit(nameSetOp, nameSetIdx)
	jumpEnd := c.emitJump(OpJump)

	c.patchJump(jumpUseApplication)
	// Application.StaticObjects(name)
	opApplication, idxApplication := c.resolveVar("Application")
	c.emit(opApplication, idxApplication)
	c.emit(OpConstant, nameConstIdx)
	c.emit(OpCallMember, staticObjectsConstIdx, 1)
	c.emit(nameSetOp, nameSetIdx)

	c.patchJump(jumpEnd)
	if useCurrentJump >= 0 {
		c.patchJumpTo(useCurrentJump, len(c.bytecode))
	}
	c.emit(nameGetOp, nameIdx)
	c.clearLastCallTarget()
	return true
}

// compileImplicitClassMemberCall compiles class-local member calls like MethodName(arg1, arg2)
// into OpMe + OpCallMember when the member belongs to the current class.
func (c *Compiler) compileImplicitClassMemberCall(name string) bool {
	if c == nil || c.currentClassName == "" || (!c.isLocal && !c.dynamicMemberResolution) {
		return false
	}
	trimmedName := strings.TrimSpace(name)
	if _, exists := c.locals.Get(name); exists && !strings.EqualFold(trimmedName, strings.TrimSpace(c.currentFunctionName)) {
		return false
	}
	lp, ok := c.next.(*vbscript.PunctuationToken)
	if !ok || lp.Type != vbscript.PunctLParen {
		return false
	}

	hasMethod := c.hasClassMethodDeclaration(c.currentClassName, name)
	hasPropertyGet := false
	if property, exists := c.getClassPropertyDeclaration(c.currentClassName, name); exists && property != nil && property.HasGet {
		hasPropertyGet = true
	}
	if strings.EqualFold(trimmedName, strings.TrimSpace(c.currentFunctionName)) {
		hasMethod = true
	}
	if !hasMethod && !hasPropertyGet {
		// In expression context, bare Name(args) inside a class should only rewrite to
		// Me.Name(args) when one class member is already known. Otherwise preserve global
		// resolution so forward page-level functions still bind like IIS/VBScript.
		return false
	}

	c.clearLastCallTarget()
	c.emit(OpMe)
	c.move() // consume '('

	argCount := 0
	if rp, ok := c.next.(*vbscript.PunctuationToken); !ok || rp.Type != vbscript.PunctRParen {
		for {
			if comma, ok := c.next.(*vbscript.PunctuationToken); ok && comma.Type == vbscript.PunctComma {
				emptyIdx := c.addConstant(NewEmpty())
				c.emit(OpConstant, emptyIdx)
			} else if rp, ok := c.next.(*vbscript.PunctuationToken); ok && rp.Type == vbscript.PunctRParen {
				emptyIdx := c.addConstant(NewEmpty())
				c.emit(OpConstant, emptyIdx)
			} else {
				argStartPos := len(c.bytecode)
				c.parseExpression(PrecNone)
				c.patchArgRefInBytecode(argStartPos)
			}
			argCount++
			if comma, ok := c.next.(*vbscript.PunctuationToken); ok && comma.Type == vbscript.PunctComma {
				c.move()
				continue
			}
			break
		}
	}

	rp, ok := c.next.(*vbscript.PunctuationToken)
	if !ok || rp.Type != vbscript.PunctRParen {
		panic(c.vbCompileError(vbscript.ExpectedRParen, "Expected ')' after class member call arguments"))
	}
	c.move() // consume ')'

	midx := c.addConstant(NewString(name))
	c.emit(OpCallMember, midx, argCount)
	return true
}

// parseExpression implements the Pratt Parser loop.
// patchArgRefInBytecode inspects the bytes emitted since startPos and replaces a simple
// OpGetGlobal, OpGetLocal, or OpGetClassMember with the corresponding OpArg*Ref opcode.
// This enables ByRef parameter write-back for simple variable arguments at call sites.
// For global slots, only user-declared slots (idx >= c.userGlobalsStart) are patched.
func (c *Compiler) patchArgRefInBytecode(startPos int) {
	emitted := len(c.bytecode) - startPos
	if emitted < 3 {
		return
	}

	// Simple ByRef candidates are emitted as OpGetGlobal/OpGetLocal/OpGetClassMember
	// (+ optional OpCoerceToValue).
	op := OpCode(c.bytecode[startPos])
	idx := int(binary.BigEndian.Uint16(c.bytecode[startPos+1:]))
	hasTrailingCoerce := emitted == 4 && OpCode(c.bytecode[startPos+3]) == OpCoerceToValue
	if !(emitted == 3 || hasTrailingCoerce) {
		return
	}

	switch op {
	case OpGetGlobal:
		// A bare zero-arg global Function name in argument position must remain
		// a value expression so OpCoerceToValue can auto-invoke it (IIS/VBScript).
		// Rewriting to OpArgGlobalRef would pass the function slot itself instead.
		if c.isGlobalZeroArgFunctionSlot(idx) {
			return
		}
		if hasTrailingCoerce {
			name := strings.ToLower(strings.TrimSpace(c.Globals.names[idx]))
			if !c.declaredGlobals[name] {
				return
			}
		}
		// Only patch user-declared variable slots; intrinsics, builtins, and VBScript constants are read-only.
		if idx >= c.userGlobalsStart {
			c.bytecode[startPos] = byte(OpArgGlobalRef)
		}
	case OpGetLocal:
		// All local variables are user-declared and eligible for ByRef write-back.
		c.bytecode[startPos] = byte(OpArgLocalRef)
	case OpGetClassMember:
		// Class member fields are mutable slots and must preserve ByRef write-back.
		c.bytecode[startPos] = byte(OpArgClassMemberRef)
	default:
		return
	}

	if hasTrailingCoerce {
		// ByRef must pass a writable slot reference, not an implicitly coerced value.
		copy(c.bytecode[startPos+3:], c.bytecode[startPos+4:])
		c.bytecode = c.bytecode[:len(c.bytecode)-1]
	}
}

// isGlobalZeroArgFunctionSlot reports whether one global slot name is a known
// zero-argument Function declaration tracked during compilation.
func (c *Compiler) isGlobalZeroArgFunctionSlot(idx int) bool {
	if c == nil || c.Globals == nil || idx < 0 || idx >= len(c.Globals.names) {
		return false
	}
	name := strings.ToLower(strings.TrimSpace(c.Globals.names[idx]))
	if name == "" {
		return false
	}
	return c.globalZeroArgFuncs[name]
}

// emitAutoCallForBareGlobalBeforeMemberAccess preserves IIS/VBScript semantics for
// bare zero-arg global Functions used as object expressions before one member chain,
// e.g. getIntranetHomePage.iId.
func (c *Compiler) emitAutoCallForBareGlobalBeforeMemberAccess() {
	if c == nil || !c.lastCallIsGlobal || c.lastCallTargetPos < 0 {
		return
	}
	if c.lastCallTargetPos+3 != len(c.bytecode) {
		return
	}
	if OpCode(c.bytecode[c.lastCallTargetPos]) != OpGetGlobal {
		return
	}
	idx := int(binary.BigEndian.Uint16(c.bytecode[c.lastCallTargetPos+1:]))
	if !c.isGlobalZeroArgFunctionSlot(idx) {
		return
	}
	c.emit(OpCoerceToValue)
}

// parseExpression implements the Pratt Parser loop.
func (c *Compiler) parseExpression(precedence Precedence) {
	token := c.move()
	prefixRule := c.getPrefixRule(token)

	if prefixRule == nil {
		panic(c.vbCompileError(vbscript.SyntaxError, fmt.Sprintf("Syntax error: Unexpected token %T in expression", token)))
	}

	prefixRule(c, token)

	for precedence < c.getPrecedence(c.next) {
		infixToken := c.move()
		infixRule := c.getInfixRule(infixToken)
		if infixRule == nil {
			break
		}
		infixRule(c, infixToken)
	}
}

// getPrecedence returns the operator precedence for the given token.
func (c *Compiler) getPrecedence(token vbscript.Token) Precedence {
	switch t := token.(type) {
	case *vbscript.PunctuationToken:
		switch t.Type {
		case vbscript.PunctDot, vbscript.PunctLParen:
			return PrecCall
		case vbscript.PunctExp:
			return PrecExp
		case vbscript.PunctAmp:
			return PrecConcat
		case vbscript.PunctPlus, vbscript.PunctMinus:
			return PrecTerm
		case vbscript.PunctStar, vbscript.PunctSlash, vbscript.PunctBackslash:
			return PrecFactor
		case vbscript.PunctEqual, vbscript.PunctNotEqual, vbscript.PunctLess, vbscript.PunctGreater, vbscript.PunctLessOrEqual, vbscript.PunctGreaterOrEqual:
			return PrecEquality
		}
	case *vbscript.KeywordToken:
		switch t.Keyword {
		case vbscript.KeywordEqv:
			return PrecEqv
		case vbscript.KeywordImp:
			return PrecImp
		case vbscript.KeywordXor:
			return PrecXor
		case vbscript.KeywordAnd:
			return PrecAnd
		case vbscript.KeywordOr:
			return PrecOr
		case vbscript.KeywordNot:
			return PrecNot
		case vbscript.KeywordMod:
			return PrecFactor
		case vbscript.KeywordIs:
			return PrecEquality
		}
	}
	return PrecNone
}

// getPrefixRule returns the rule for tokens at the start of an expression.
func (c *Compiler) getPrefixRule(token vbscript.Token) func(*Compiler, vbscript.Token) {
	switch t := token.(type) {
	case *vbscript.TrueLiteralToken:
		return func(c *Compiler, tk vbscript.Token) {
			c.clearLastCallTarget()
			idx := c.addConstant(NewBool(true))
			c.emit(OpConstant, idx)
		}
	case *vbscript.FalseLiteralToken:
		return func(c *Compiler, tk vbscript.Token) {
			c.clearLastCallTarget()
			idx := c.addConstant(NewBool(false))
			c.emit(OpConstant, idx)
		}
	case *vbscript.NullLiteralToken:
		return func(c *Compiler, tk vbscript.Token) {
			c.clearLastCallTarget()
			idx := c.addConstant(Value{Type: VTNull})
			c.emit(OpConstant, idx)
		}
	case *vbscript.EmptyLiteralToken:
		return func(c *Compiler, tk vbscript.Token) {
			c.clearLastCallTarget()
			idx := c.addConstant(Value{Type: VTEmpty})
			c.emit(OpConstant, idx)
		}
	case *vbscript.NothingLiteralToken:
		return func(c *Compiler, tk vbscript.Token) {
			c.clearLastCallTarget()
			// Nothing is a null object reference for Set/Is semantics.
			idx := c.addConstant(Value{Type: VTObject, Num: 0})
			c.emit(OpConstant, idx)
		}
	case *vbscript.DateLiteralToken:
		return func(c *Compiler, tk vbscript.Token) {
			c.clearLastCallTarget()
			idx := c.addConstant(NewDate(t.Value))
			c.emit(OpConstant, idx)
		}
	case *vbscript.DecIntegerLiteralToken:
		return func(c *Compiler, tk vbscript.Token) {
			c.clearLastCallTarget()
			idx := c.addConstant(NewInteger(t.Value))
			c.emit(OpConstant, idx)
		}
	case *vbscript.HexIntegerLiteralToken:
		return func(c *Compiler, tk vbscript.Token) {
			c.clearLastCallTarget()
			idx := c.addConstant(NewInteger(t.Value))
			c.emit(OpConstant, idx)
		}
	case *vbscript.OctIntegerLiteralToken:
		return func(c *Compiler, tk vbscript.Token) {
			c.clearLastCallTarget()
			idx := c.addConstant(NewInteger(t.Value))
			c.emit(OpConstant, idx)
		}
	case *vbscript.FloatLiteralToken:
		return func(c *Compiler, tk vbscript.Token) {
			c.clearLastCallTarget()
			idx := c.addConstant(NewDouble(t.Value))
			c.emit(OpConstant, idx)
		}
	case *vbscript.StringLiteralToken:
		return func(c *Compiler, tk vbscript.Token) {
			c.clearLastCallTarget()
			idx := c.addConstant(NewString(t.Value))
			c.emit(OpConstant, idx)
		}
	case *vbscript.IdentifierToken:
		return func(c *Compiler, tk vbscript.Token) {
			if c.compileImplicitClassMemberCall(t.Name) {
				return
			}
			c.emitIdentifierValue(t.Name)
		}
	case *vbscript.KeywordOrIdentifierToken:
		return func(c *Compiler, tk vbscript.Token) {
			if c.compileImplicitClassMemberCall(t.Name) {
				return
			}
			c.emitIdentifierValue(t.Name)
		}
	case *vbscript.ExtendedIdentifierToken:
		return func(c *Compiler, tk vbscript.Token) {
			name := t.Name
			if len(name) >= 2 && name[0] == '[' && name[len(name)-1] == ']' {
				name = name[1 : len(name)-1]
			}
			if c.compileImplicitClassMemberCall(name) {
				return
			}
			c.emitIdentifierValue(name)
		}
	case *vbscript.PunctuationToken:
		if t.Type == vbscript.PunctLParen {
			return func(c *Compiler, tk vbscript.Token) {
				c.clearLastCallTarget()
				c.parseExpression(PrecNone)
				if p, ok := c.next.(*vbscript.PunctuationToken); !ok || p.Type != vbscript.PunctRParen {
					panic(c.vbSyntaxError(vbscript.ExpectedRParen))
				}
				c.move() // consume ')'
			}
		} else if t.Type == vbscript.PunctMinus {
			return func(c *Compiler, tk vbscript.Token) {
				c.clearLastCallTarget()
				c.parseExpression(PrecUnary)
				c.emit(OpNeg)
			}
		} else if t.Type == vbscript.PunctPlus {
			return func(c *Compiler, tk vbscript.Token) {
				c.clearLastCallTarget()
				c.parseExpression(PrecUnary)
			}
		} else if t.Type == vbscript.PunctDot && c.withDepth > 0 {
			// Expression-level '.Member' inside a With block: x = .Prop, f(.Method(a))
			return func(c *Compiler, tk vbscript.Token) {
				c.clearLastCallTarget()
				// Push the innermost With-subject object.
				c.emit(OpWithLoad)
				name := c.expectIdentifier()
				midx := c.addConstant(NewString(name))
				if lp, ok := c.next.(*vbscript.PunctuationToken); ok && lp.Type == vbscript.PunctLParen {
					// .Method(args) — OpCallMember reads member from inline bytecode operand.
					c.move() // consume '('
					argCount := 0
					if rp, ok2 := c.next.(*vbscript.PunctuationToken); !ok2 || rp.Type != vbscript.PunctRParen {
						for {
							if comma, ok3 := c.next.(*vbscript.PunctuationToken); ok3 && comma.Type == vbscript.PunctComma {
								emptyIdx := c.addConstant(NewEmpty())
								c.emit(OpConstant, emptyIdx)
							} else if rp2, ok3 := c.next.(*vbscript.PunctuationToken); ok3 && rp2.Type == vbscript.PunctRParen {
								emptyIdx := c.addConstant(NewEmpty())
								c.emit(OpConstant, emptyIdx)
							} else {
								argStartPos := len(c.bytecode)
								c.parseExpression(PrecNone)
								c.patchArgRefInBytecode(argStartPos)
							}
							argCount++
							if comma, ok3 := c.next.(*vbscript.PunctuationToken); ok3 && comma.Type == vbscript.PunctComma {
								c.move()
							} else {
								break
							}
						}
					}
					if rp, ok2 := c.next.(*vbscript.PunctuationToken); !ok2 || rp.Type != vbscript.PunctRParen {
						panic(c.vbCompileError(vbscript.ExpectedRParen, "Expected ')' after member call arguments"))
					}
					c.move() // consume ')'
					c.emit(OpCallMember, midx, argCount)
					c.emitTrailingCoerceIfValueContext()
				} else {
					// .Prop — OpMemberGet reads target and member-name both from the stack.
					// Push member name as constant, then emit the 0-operand OpMemberGet.
					c.emit(OpConstant, midx)
					c.emit(OpMemberGet)
					c.emitTrailingCoerceIfValueContext()
				}
			}
		}
	case *vbscript.KeywordToken:
		if t.Keyword == vbscript.KeywordNot {
			return func(c *Compiler, tk vbscript.Token) {
				c.clearLastCallTarget()
				// VBScript evaluates "Not obj Is Nothing" as "Not (obj Is Nothing)",
				// so Not must parse with PrecNot to include Is/Is Not in its operand.
				c.parseExpression(PrecNot)
				c.emit(OpNot)
			}
		}
		if t.Keyword == vbscript.KeywordNew {
			return func(c *Compiler, tk vbscript.Token) {
				c.clearLastCallTarget()
				className := c.expectIdentifier()
				classNameIdx := c.addConstant(NewString(className))
				c.emit(OpNewClass, classNameIdx)
			}
		}
		// Me refers to the current class instance inside a method/property.
		if t.Keyword == vbscript.KeywordMe {
			return func(c *Compiler, tk vbscript.Token) {
				c.clearLastCallTarget()
				c.emit(OpMe)
			}
		}
	}
	return nil
}

// getInfixRule returns the rule for operators between expressions.
func (c *Compiler) getInfixRule(token vbscript.Token) func(*Compiler, vbscript.Token) {
	switch t := token.(type) {
	case *vbscript.PunctuationToken:
		switch t.Type {
		case vbscript.PunctDot:
			return func(c *Compiler, tk vbscript.Token) {
				c.undoTrailingCoerce() // member access needs raw object reference
				c.emitAutoCallForBareGlobalBeforeMemberAccess()
				c.clearLastCallTarget()
				name := c.expectIdentifier()
				// If followed by '(', compile as a direct member call (OpCallMember).
				// This avoids the OpMemberGet+OpCall pattern that breaks zero-arg
				// getter calls and multi-arg collections (e.g. Cookies("k","Domain")).
				if lp, ok := c.next.(*vbscript.PunctuationToken); ok && lp.Type == vbscript.PunctLParen {
					c.move() // consume '('
					argCount := 0
					if rp, ok2 := c.next.(*vbscript.PunctuationToken); !ok2 || rp.Type != vbscript.PunctRParen {
						for {
							if comma, ok3 := c.next.(*vbscript.PunctuationToken); ok3 && comma.Type == vbscript.PunctComma {
								emptyIdx := c.addConstant(NewEmpty())
								c.emit(OpConstant, emptyIdx)
							} else if rp, ok3 := c.next.(*vbscript.PunctuationToken); ok3 && rp.Type == vbscript.PunctRParen {
								emptyIdx := c.addConstant(NewEmpty())
								c.emit(OpConstant, emptyIdx)
							} else {
								argStartPos := len(c.bytecode)
								callArgStartPos := len(c.bytecode)
								c.parseExpression(PrecNone)
								c.patchArgRefInBytecode(callArgStartPos)
								c.patchArgRefInBytecode(argStartPos)
							}
							argCount++
							if comma, ok3 := c.next.(*vbscript.PunctuationToken); ok3 && comma.Type == vbscript.PunctComma {
								c.move()
							} else {
								break
							}
						}
					}
					if rp, ok2 := c.next.(*vbscript.PunctuationToken); !ok2 || rp.Type != vbscript.PunctRParen {
						panic(c.vbCompileError(vbscript.ExpectedRParen, "Expected ')' after member call arguments"))
					}
					c.move() // consume ')'
					midx := c.addConstant(NewString(name))
					c.emit(OpCallMember, midx, argCount)
					c.emitTrailingCoerceIfValueContext()
				} else {
					// Plain property access: Obj.Prop — emit OpMemberGet.
					// Chaining (Obj.A.B) is handled naturally by repeated Pratt infix application.
					idx := c.addConstant(NewString(name))
					c.emit(OpConstant, idx)
					c.emit(OpMemberGet)
					c.emitTrailingCoerceIfValueContext()
				}
			}
		case vbscript.PunctLParen:
			return func(c *Compiler, tk vbscript.Token) {
				c.undoTrailingCoerce() // function/method call needs raw callable reference
				callTargetName := c.lastCallTargetName
				callTargetPos := c.lastCallTargetPos
				callTargetIsGlobal := c.lastCallIsGlobal
				argCount := 0
				argExprStart := len(c.bytecode)
				// Check for immediate closing paren (empty args)
				if p, ok := c.next.(*vbscript.PunctuationToken); !ok || p.Type != vbscript.PunctRParen {
					for {
						if comma, ok := c.next.(*vbscript.PunctuationToken); ok && comma.Type == vbscript.PunctComma {
							emptyIdx := c.addConstant(NewEmpty())
							c.emit(OpConstant, emptyIdx)
						} else if rp, ok := c.next.(*vbscript.PunctuationToken); ok && rp.Type == vbscript.PunctRParen {
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
				if p, ok := c.next.(*vbscript.PunctuationToken); !ok || p.Type != vbscript.PunctRParen {
					panic(c.vbCompileError(vbscript.ExpectedRParen, "Expected ')' after function arguments"))
				}
				c.move() // consume ')'

				if c.tryInlineEvalCall(callTargetName, callTargetPos, argExprStart, argCount) {
					return
				}

				c.emit(OpCall, argCount)
				if callTargetIsGlobal {
					c.registerForwardCallPatch(callTargetName, callTargetPos)
				}
				c.emitTrailingCoerceIfValueContext()
				c.clearLastCallTarget()
			}
		case vbscript.PunctAmp:
			return func(c *Compiler, tk vbscript.Token) {
				c.clearLastCallTarget()
				c.parseExpression(PrecConcat)
				c.emit(OpConcat)
			}
		case vbscript.PunctExp:
			return func(c *Compiler, tk vbscript.Token) {
				c.clearLastCallTarget()
				c.parseExpression(PrecExp - 1)
				c.emit(OpPow)
			}
		case vbscript.PunctEqual:
			return func(c *Compiler, tk vbscript.Token) {
				c.clearLastCallTarget()
				c.parseExpression(PrecEquality)
				c.emit(OpEq)
			}
		case vbscript.PunctNotEqual:
			return func(c *Compiler, tk vbscript.Token) {
				c.clearLastCallTarget()
				c.parseExpression(PrecEquality)
				c.emit(OpNeq)
			}
		case vbscript.PunctLess:
			return func(c *Compiler, tk vbscript.Token) {
				c.clearLastCallTarget()
				c.parseExpression(PrecEquality)
				c.emit(OpLt)
			}
		case vbscript.PunctGreater:
			return func(c *Compiler, tk vbscript.Token) {
				c.clearLastCallTarget()
				c.parseExpression(PrecEquality)
				c.emit(OpGt)
			}
		case vbscript.PunctLessOrEqual:
			return func(c *Compiler, tk vbscript.Token) {
				c.clearLastCallTarget()
				c.parseExpression(PrecEquality)
				c.emit(OpLte)
			}
		case vbscript.PunctGreaterOrEqual:
			return func(c *Compiler, tk vbscript.Token) {
				c.clearLastCallTarget()
				c.parseExpression(PrecEquality)
				c.emit(OpGte)
			}
		case vbscript.PunctPlus:
			return func(c *Compiler, tk vbscript.Token) {
				c.clearLastCallTarget()
				c.parseExpression(PrecTerm)
				c.emit(OpAdd)
			}
		case vbscript.PunctMinus:
			return func(c *Compiler, tk vbscript.Token) {
				c.clearLastCallTarget()
				c.parseExpression(PrecTerm)
				c.emit(OpSub)
			}
		case vbscript.PunctStar:
			return func(c *Compiler, tk vbscript.Token) {
				c.clearLastCallTarget()
				c.parseExpression(PrecFactor)
				c.emit(OpMul)
			}
		case vbscript.PunctSlash:
			return func(c *Compiler, tk vbscript.Token) {
				c.clearLastCallTarget()
				c.parseExpression(PrecFactor)
				c.emit(OpDiv)
			}
		case vbscript.PunctBackslash:
			return func(c *Compiler, tk vbscript.Token) {
				c.clearLastCallTarget()
				c.parseExpression(PrecFactor)
				c.emit(OpIDiv)
			}
		}
	case *vbscript.KeywordToken:
		switch t.Keyword {
		case vbscript.KeywordMod:
			return func(c *Compiler, tk vbscript.Token) {
				c.clearLastCallTarget()
				c.parseExpression(PrecFactor)
				c.emit(OpMod)
			}
		case vbscript.KeywordIs:
			return func(c *Compiler, tk vbscript.Token) {
				c.undoTrailingCoerce() // Is/Is Not compare raw object references
				isNot := false
				if nextToken, ok := c.next.(*vbscript.KeywordToken); ok && nextToken.Keyword == vbscript.KeywordNot {
					c.move()
					isNot = true
				}
				c.parseExpression(PrecEquality)
				if isNot {
					c.emit(OpIsNotRef)
				} else {
					c.emit(OpIsRef)
				}
			}
		case vbscript.KeywordAnd:
			return func(c *Compiler, tk vbscript.Token) {
				c.clearLastCallTarget()
				c.parseExpression(PrecAnd)
				c.emit(OpAnd)
			}
		case vbscript.KeywordOr:
			return func(c *Compiler, tk vbscript.Token) {
				c.clearLastCallTarget()
				c.parseExpression(PrecOr)
				c.emit(OpOr)
			}
		case vbscript.KeywordXor:
			return func(c *Compiler, tk vbscript.Token) {
				c.clearLastCallTarget()
				c.parseExpression(PrecXor)
				c.emit(OpXor)
			}
		case vbscript.KeywordEqv:
			return func(c *Compiler, tk vbscript.Token) {
				c.clearLastCallTarget()
				c.parseExpression(PrecEqv)
				c.emit(OpEqv)
			}
		case vbscript.KeywordImp:
			return func(c *Compiler, tk vbscript.Token) {
				c.clearLastCallTarget()
				c.parseExpression(PrecImp)
				c.emit(OpImp)
			}
		}
	}
	return nil
}

// Compile translates the entire source into bytecode.
func (c *Compiler) Compile() (err error) {
	defer func() {
		// Dump the preprocessed source on every exit path, including panics.
		// dumpPreprocessedSource respects the SetDumpPreprocessedSourceEnabled flag
		// and the DUMP_PREPROCESSED_SOURCE env var; it is a no-op when both are off.
		dumpPreprocessedSource(c.sourceCode, c.sourceName)

		if r := recover(); r != nil {
			switch recovered := r.(type) {
			case error:
				err = c.normalizeCompileError(recovered)
			default:
				err = c.normalizeCompileError(fmt.Errorf("%v", recovered))
			}
		}
	}()

	if c.lexerMode == vbscript.ModeASP {
		c.includeDeps = c.includeDeps[:0]
		if strings.Contains(strings.ToLower(c.sourceCode), "#include") {
			expanded, mappedLines, preprocessErr := preprocessASPIncludesWithDeps(c.sourceCode, c.sourceName, map[string]bool{}, 0, &c.includeDeps)
			if preprocessErr != nil {
				return c.normalizeCompileError(c.vbCompileError(vbscript.SyntaxError, preprocessErr.Error()))
			}
			if expanded != c.sourceCode {
				c.sourceCode = expanded
			}
			c.lineMap = mappedLines
		}

		// Remove known empty ASP blocks before metadata scanning and lexing.
		cleanedASP := stripEmptyASPBlocks(c.sourceCode)
		if cleanedASP != c.sourceCode {
			c.sourceCode = cleanedASP
		}

		libs := detectMetadataLibraries(c.sourceCode)
		injected := c.injectTypeLibConstants(getMetadataLibraryConstants(libs))
		c.emitInjectedConstantInitializers(injected)

		// Strip <!-- METADATA TYPE="TypeLib" --> directives from the source so they
		// do not appear in rendered HTML output, matching IIS behaviour.
		stripped := stripMetadataDirectives(c.sourceCode)
		if stripped != c.sourceCode {
			c.sourceCode = stripped
		}
	}

	c.resetTokenStream()
	if c.isEval {
		c.parseExpression(PrecNone)
		c.emit(OpHalt)
		return nil
	}

	compiledDefinitionBounds := c.compileDefinitionPreBindingPass()

	c.resetTokenStream()
	skipDefinitionStarts := make(map[int]int, len(compiledDefinitionBounds))
	for i := range compiledDefinitionBounds {
		skipDefinitionStarts[compiledDefinitionBounds[i].start] = compiledDefinitionBounds[i].end
	}

	for !c.matchEof() {
		if end, shouldSkip := skipDefinitionStarts[c.tokenIndex]; shouldSkip {
			c.skipDefinitionBlock(end)
			continue
		}

		switch t := c.next.(type) {
		case *vbscript.HTMLToken:
			idx := c.addConstant(NewString(t.Content))
			c.emit(OpWriteStatic, idx)
			c.move()
		case *vbscript.ASPExpressionStartToken:
			c.move()
			c.emitCurrentDebugLocation()
			c.parseExpression(PrecNone)
			c.emit(OpWrite)
		case *vbscript.ASPDirectiveStartToken:
			c.move()
			c.compileASPDirective()
		case *vbscript.ASPIncludeToken:
			c.move()
		case *vbscript.ASPObjectToken:
			c.move()
			c.compileASPObjectDeclaration(t)
		case *vbscript.ASPCodeStartToken, *vbscript.ASPCodeEndToken:
			c.move()
		case *vbscript.LineTerminationToken, *vbscript.ColonLineTerminationToken, *vbscript.CommentToken:
			c.move()
		default:
			c.emitCurrentDebugLocation()
			c.parseStatement()
		}
	}
	c.emit(OpHalt)
	return nil
}
