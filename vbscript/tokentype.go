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

// TokenType represents the type of a token
type TokenType int

const (
	TokenEOF TokenType = iota
	TokenLineTermination
	TokenComment
	TokenStringLiteral
	TokenDecIntegerLiteral
	TokenHexIntegerLiteral
	TokenOctIntegerLiteral
	TokenDateLiteral
	TokenFloatLiteral
	TokenTrueLiteral
	TokenFalseLiteral
	TokenNullLiteral
	TokenEmptyLiteral
	TokenNothingLiteral
	TokenIdentifier
)

// Punctuation represents punctuation token types
type Punctuation int

const (
	PunctLParen Punctuation = iota
	PunctRParen
	PunctLBracket
	PunctRBracket
	PunctDot
	PunctComma
	PunctPlus
	PunctMinus
	PunctSlash
	PunctBackslash
	PunctStar
	PunctAmp
	PunctExp
	PunctEqual
	PunctNotEqual
	PunctLess
	PunctLessOrEqual
	PunctGreater
	PunctGreaterOrEqual
)

// Keyword represents VBScript keywords
type Keyword int

const (
	KeywordStep Keyword = iota
	KeywordProperty
	KeywordExplicit
	KeywordError
	KeywordErase
	KeywordDefault
	KeywordAnd
	KeywordByRef
	KeywordByVal
	KeywordCall
	KeywordCase
	KeywordClass
	KeywordConst
	KeywordDim
	KeywordDo
	KeywordEach
	KeywordElse
	KeywordElseIf
	KeywordEnd
	KeywordEqv
	KeywordExit
	KeywordFor
	KeywordFunction
	KeywordGet
	KeywordGoto
	KeywordIf
	KeywordImp
	KeywordIn
	KeywordXor
	KeywordWith
	KeywordWhile
	KeywordWEnd
	KeywordTo
	KeywordUntil
	KeywordThen
	KeywordSub
	KeywordSet
	KeywordSelect
	KeywordResume
	KeywordReDim
	KeywordPublic
	KeywordPrivate
	KeywordPreserve
	KeywordOr
	KeywordOption
	KeywordOn
	KeywordNot
	KeywordNext
	KeywordNew
	KeywordMod
	KeywordLoop
	KeywordLet
	KeywordIs
	KeywordBinary
	KeywordCompare
	KeywordText
	KeywordBase
)

// String returns the string representation of a Keyword
func (k Keyword) String() string {
	switch k {
	case KeywordStep:
		return "Step"
	case KeywordProperty:
		return "Property"
	case KeywordExplicit:
		return "Explicit"
	case KeywordError:
		return "Error"
	case KeywordErase:
		return "Erase"
	case KeywordDefault:
		return "Default"
	case KeywordAnd:
		return "And"
	case KeywordByRef:
		return "ByRef"
	case KeywordByVal:
		return "ByVal"
	case KeywordCall:
		return "Call"
	case KeywordCase:
		return "Case"
	case KeywordClass:
		return "Class"
	case KeywordConst:
		return "Const"
	case KeywordDim:
		return "Dim"
	case KeywordDo:
		return "Do"
	case KeywordEach:
		return "Each"
	case KeywordElse:
		return "Else"
	case KeywordElseIf:
		return "ElseIf"
	case KeywordEnd:
		return "End"
	case KeywordEqv:
		return "Eqv"
	case KeywordExit:
		return "Exit"
	case KeywordFor:
		return "For"
	case KeywordFunction:
		return "Function"
	case KeywordGet:
		return "Get"
	case KeywordGoto:
		return "Goto"
	case KeywordIf:
		return "If"
	case KeywordImp:
		return "Imp"
	case KeywordIn:
		return "In"
	case KeywordXor:
		return "Xor"
	case KeywordWith:
		return "With"
	case KeywordWhile:
		return "While"
	case KeywordWEnd:
		return "WEnd"
	case KeywordTo:
		return "To"
	case KeywordUntil:
		return "Until"
	case KeywordThen:
		return "Then"
	case KeywordSub:
		return "Sub"
	case KeywordSet:
		return "Set"
	case KeywordSelect:
		return "Select"
	case KeywordResume:
		return "Resume"
	case KeywordReDim:
		return "ReDim"
	case KeywordPublic:
		return "Public"
	case KeywordPrivate:
		return "Private"
	case KeywordPreserve:
		return "Preserve"
	case KeywordOr:
		return "Or"
	case KeywordOption:
		return "Option"
	case KeywordOn:
		return "On"
	case KeywordNot:
		return "Not"
	case KeywordNext:
		return "Next"
	case KeywordNew:
		return "New"
	case KeywordMod:
		return "Mod"
	case KeywordLoop:
		return "Loop"
	case KeywordLet:
		return "Let"
	case KeywordIs:
		return "Is"
	case KeywordBinary:
		return "Binary"
	case KeywordCompare:
		return "Compare"
	case KeywordText:
		return "Text"
	case KeywordBase:
		return "Base"
	default:
		return "Unknown"
	}
}
