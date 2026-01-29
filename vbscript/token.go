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

import "time"

// Token is the base interface for all token types
type Token interface {
	GetStart() int
	SetStart(int)
	GetEnd() int
	SetEnd(int)
	GetLineNumber() int
	SetLineNumber(int)
	GetLineStart() int
	SetLineStart(int)
}

// BaseToken is the common base for all tokens
type BaseToken struct {
	Start      int
	End        int
	LineNumber int
	LineStart  int
}

// GetStart returns the start position
func (t *BaseToken) GetStart() int { return t.Start }

// SetStart sets the start position
func (t *BaseToken) SetStart(pos int) { t.Start = pos }

// GetEnd returns the end position
func (t *BaseToken) GetEnd() int { return t.End }

// SetEnd sets the end position
func (t *BaseToken) SetEnd(pos int) { t.End = pos }

// GetLineNumber returns the line number
func (t *BaseToken) GetLineNumber() int { return t.LineNumber }

// SetLineNumber sets the line number
func (t *BaseToken) SetLineNumber(line int) { t.LineNumber = line }

// GetLineStart returns the line start position
func (t *BaseToken) GetLineStart() int { return t.LineStart }

// SetLineStart sets the line start position
func (t *BaseToken) SetLineStart(pos int) { t.LineStart = pos }

// EOFToken represents the end of file
type EOFToken struct {
	BaseToken
}

// LineTerminationToken represents a line termination
type LineTerminationToken struct {
	BaseToken
}

// ColonLineTerminationToken represents a colon line termination (:)
type ColonLineTerminationToken struct {
	LineTerminationToken
}

// CommentToken represents a comment
type CommentToken struct {
	BaseToken
	Comment string // The comment text (without delimiters)
	IsRem   bool   // true if comment starts with 'REM', false if starts with '''
}

// LiteralToken is the base for literal tokens
type LiteralToken struct {
	BaseToken
}

// StringLiteralToken represents a string literal
type StringLiteralToken struct {
	LiteralToken
	Value string
}

// DecIntegerLiteralToken represents a decimal integer literal
type DecIntegerLiteralToken struct {
	LiteralToken
	Value int64
}

// HexIntegerLiteralToken represents a hexadecimal integer literal
type HexIntegerLiteralToken struct {
	DecIntegerLiteralToken
}

// OctIntegerLiteralToken represents an octal integer literal
type OctIntegerLiteralToken struct {
	DecIntegerLiteralToken
}

// DateLiteralToken represents a date literal
type DateLiteralToken struct {
	LiteralToken
	Value time.Time
}

// FloatLiteralToken represents a floating-point literal
type FloatLiteralToken struct {
	LiteralToken
	Value float64
}

// TrueLiteralToken represents the 'True' keyword
type TrueLiteralToken struct {
	LiteralToken
}

// FalseLiteralToken represents the 'False' keyword
type FalseLiteralToken struct {
	LiteralToken
}

// NullLiteralToken represents the 'Null' keyword
type NullLiteralToken struct {
	LiteralToken
}

// NothingLiteralToken represents the 'Nothing' keyword
type NothingLiteralToken struct {
	LiteralToken
}

// EmptyLiteralToken represents the 'Empty' keyword
type EmptyLiteralToken struct {
	LiteralToken
}

// IdentifierToken represents an identifier
type IdentifierToken struct {
	BaseToken
	Name string
}

// String returns the identifier name
func (t *IdentifierToken) String() string {
	return t.Name
}

// KeywordToken represents a keyword token
type KeywordToken struct {
	BaseToken
	Keyword Keyword
	Name    string
}

// String returns the keyword name
func (t *KeywordToken) String() string {
	return t.Name
}

// KeywordOrIdentifierToken represents a token that could be either a keyword or identifier
type KeywordOrIdentifierToken struct {
	BaseToken
	Name    string
	Keyword Keyword
}

// String returns the token name
func (t *KeywordOrIdentifierToken) String() string {
	return t.Name
}

// ExtendedIdentifierToken represents an extended identifier (enclosed in brackets)
type ExtendedIdentifierToken struct {
	IdentifierToken
}

// String returns the extended identifier with brackets
func (t *ExtendedIdentifierToken) String() string {
	return "[" + t.Name + "]"
}

// PunctuationToken represents a punctuation token
type PunctuationToken struct {
	BaseToken
	Type Punctuation
}

// InvalidToken represents an invalid token
type InvalidToken struct {
	BaseToken
}
