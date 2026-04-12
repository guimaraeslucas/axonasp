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

import (
	"strconv"
	"strings"
)

// LexerMode represents the mode of the lexer
type LexerMode int

const (
	ModeVBScript LexerMode = iota
	ModeASP
	ModeEval
)

// ASPBlockType represents the type of ASP block
type ASPBlockType int

const (
	BlockTypeNone ASPBlockType = iota
	BlockTypePercent
	BlockTypeScript
)

// Lexer represents a VBScript lexical analyzer
type Lexer struct {
	Code              string
	Index             int
	CurrentLine       int
	CurrentLineStart  int
	Length            int
	sb                strings.Builder
	runes             []rune
	asciiOnly         bool
	Mode              LexerMode
	InASPBlock        bool
	BlockType         ASPBlockType
	skipHTMLLeadingNL bool
}

func isASCIIString(value string) bool {
	for i := 0; i < len(value); i++ {
		if value[i] >= 0x80 {
			return false
		}
	}
	return true
}

// NewLexer creates a new Lexer instance
func NewLexer(code string) *Lexer {
	asciiOnly := isASCIIString(code)
	var r []rune
	length := len(code)
	if !asciiOnly {
		r = []rune(code)
		length = len(r)
	}
	if code == "" {
		return &Lexer{
			Code:             code,
			Index:            0,
			CurrentLine:      0,
			CurrentLineStart: 0,
			Length:           0,
			runes:            r,
			asciiOnly:        asciiOnly,
		}
	}
	return &Lexer{
		Code:             code,
		Index:            0,
		CurrentLine:      1,
		CurrentLineStart: 0,
		Length:           length,
		runes:            r,
		asciiOnly:        asciiOnly,
	}
}

// LineIndex returns the column position in the current line
func (l *Lexer) LineIndex() int {
	return l.Index - l.CurrentLineStart
}

// NextToken returns the next token from the source code
func (l *Lexer) NextToken() Token {
	if l.Mode == ModeASP && !l.InASPBlock {
		return l.nextHTML()
	}

	l.skipWhitespaces()

	if l.isEof() {
		return &EOFToken{
			BaseToken: BaseToken{
				Start:      l.Index,
				End:        l.Index,
				LineNumber: l.CurrentLine,
				LineStart:  l.CurrentLineStart,
			},
		}
	}

	c := l.getChar(l.Index)
	next := l.getChar(l.Index + 1)

	if l.Mode == ModeASP && l.InASPBlock {
		if l.BlockType == BlockTypePercent && c == '%' && next == '>' {
			start := l.Index
			l.Index += 2
			l.InASPBlock = false
			l.BlockType = BlockTypeNone
			l.skipHTMLLeadingNL = true
			return &ASPCodeEndToken{
				BaseToken: BaseToken{
					Start:      start,
					End:        l.Index,
					LineNumber: l.CurrentLine,
					LineStart:  l.CurrentLineStart,
				},
			}
		} else if l.BlockType == BlockTypeScript {
			if length, ok := l.isScriptEnd(); ok {
				start := l.Index
				l.Index += length
				l.InASPBlock = false
				l.BlockType = BlockTypeNone
				l.skipHTMLLeadingNL = true
				return &ASPCodeEndToken{
					BaseToken: BaseToken{
						Start:      start,
						End:        l.Index,
						LineNumber: l.CurrentLine,
						LineStart:  l.CurrentLineStart,
					},
				}
			}
		}
	}

	if IsLineTerminator(c) {
		return l.nextLineTermination()
	}

	comment := l.nextComment()
	if comment != nil {
		return comment
	}

	if IsIdentifierStart(c) {
		return l.nextIdentifier()
	}

	if c == '"' {
		return l.nextStringLiteral()
	}

	if c == '.' {
		if IsDecDigit(next) {
			return l.nextNumericLiteral()
		}
		return l.nextPunctuation()
	}

	if IsDecDigit(c) {
		return l.nextNumericLiteral()
	}

	if c == '&' {
		// Classic ASP allows optional whitespace between & and the hex/oct prefix:
		// "& h22" is equivalent to "&h22". Peek past spaces to find h/o.
		// Guard: only treat as a hex/oct literal if the prefix is NOT part of an identifier.
		// Example: "& Hex(x)" must NOT be treated as "&He" hex literal — 'e' is a hex digit but
		// 'x' is an identifier continuation, so "Hex" is an identifier, not a literal prefix.
		peekIdx := l.Index + 1
		for l.getChar(peekIdx) == ' ' || l.getChar(peekIdx) == '\t' {
			peekIdx++
		}
		peeked := l.getChar(peekIdx)
		if CharEquals(peeked, 'h') {
			digitCh := l.getChar(peekIdx + 1)
			if IsHexDigit(digitCh) {
				// Scan past all consecutive hex digits to check the terminator.
				// If the terminator is an alpha or underscore, this is an identifier (e.g. &Hex),
				// not a bona-fide hex literal. Only lex as numeric when the boundary is clean.
				scanIdx := peekIdx + 1
				for IsHexDigit(l.getChar(scanIdx)) {
					scanIdx++
				}
				boundary := l.getChar(scanIdx)
				if IsIdentifierStart(boundary) || boundary == '_' {
					return l.nextPunctuation()
				}
				return l.nextNumericLiteral()
			}
			return l.nextPunctuation()
		}
		if CharEquals(peeked, 'o') {
			digitCh := l.getChar(peekIdx + 1)
			if IsOctDigit(digitCh) {
				// Same guard for octal literals: reject if hex-digit run is followed by identifier chars.
				scanIdx := peekIdx + 1
				for IsOctDigit(l.getChar(scanIdx)) {
					scanIdx++
				}
				boundary := l.getChar(scanIdx)
				if IsIdentifierStart(boundary) || boundary == '_' {
					return l.nextPunctuation()
				}
				return l.nextNumericLiteral()
			}
			return l.nextPunctuation()
		}
		if IsDecDigit(next) {
			return l.nextNumericLiteral()
		}
		return l.nextPunctuation()
	}

	if c == '#' {
		return l.nextDateLiteral()
	}

	if c == '[' {
		return l.nextExtendedIdentifier()
	}

	return l.nextPunctuation()
}

// AsSequence returns all tokens as a slice
func (l *Lexer) AsSequence() []Token {
	var tokens []Token
	for !l.isEof() {
		tokens = append(tokens, l.NextToken())
	}
	return tokens
}

// Reset resets the lexer to the beginning
func (l *Lexer) Reset() {
	l.Index = 0
	if l.Length == 0 {
		l.CurrentLine = 0
	} else {
		l.CurrentLine = 1
	}
	l.CurrentLineStart = 0
	l.sb.Reset()
}

// Private helper methods

func (l *Lexer) getChar(pos int) rune {
	if pos < 0 || pos >= l.Length {
		return rune(0)
	}
	if l.asciiOnly {
		return rune(l.Code[pos])
	}
	return l.runes[pos]
}

func (l *Lexer) sliceString(start int, end int) string {
	if start < 0 {
		start = 0
	}
	if end < start {
		end = start
	}
	if end > l.Length {
		end = l.Length
	}
	if l.asciiOnly {
		return l.Code[start:end]
	}
	return string(l.runes[start:end])
}

func (l *Lexer) isEof() bool {
	return l.Index >= l.Length
}

func (l *Lexer) skipWhitespaces() {
	for !l.isEof() {
		l.skipWSOnly()

		c := l.getChar(l.Index)
		if c != '_' {
			break
		}

		// VBScript line continuation: optional horizontal whitespace, underscore,
		// optional horizontal whitespace, then newline.
		l.Index++
		l.skipWSOnly()

		if !IsNewLine(l.getChar(l.Index)) {
			panic(l.vbSyntaxError(InvalidCharacter))
		}

		l.skipNewline()
	}
}

func (l *Lexer) skipWSOnly() {
	c := l.getChar(l.Index)
	for IsWhiteSpace(c) {
		l.Index++
		c = l.getChar(l.Index)
	}
}

func (l *Lexer) skipNewline() {
	c := l.getChar(l.Index)
	l.Index++
	if IsNewLine(c) {
		if c == '\r' && l.getChar(l.Index) == '\n' {
			l.Index++
		}
		l.CurrentLine++
		l.CurrentLineStart = l.Index
	}
}

func (l *Lexer) nextExtendedIdentifier() Token {
	start := l.Index
	l.Index++ // skip '['
	nameStart := l.Index

	for !l.isEof() {
		c := l.getChar(l.Index)
		if c == ']' {
			break
		}
		if IsNewLine(c) {
			break
		}
		if c == 0 {
			break
		}
		l.Index++
	}

	if l.getChar(l.Index) != ']' {
		panic(l.vbSyntaxError(ExpectedRBracket))
	}

	name := l.sliceString(nameStart, l.Index)

	l.Index++

	return &IdentifierToken{
		BaseToken: BaseToken{
			Start:      start,
			End:        l.Index,
			LineNumber: l.CurrentLine,
			LineStart:  l.CurrentLineStart,
		},
		Name: name,
	}

}

func (l *Lexer) nextDateLiteral() Token {
	start := l.Index
	l.Index++ // skip '#'
	l.sb.Reset()

	for !l.isEof() {
		c := l.getChar(l.Index)
		if c == '#' || IsNewLine(c) {
			break
		}
		l.sb.WriteRune(c)
		l.Index++
	}

	if l.getChar(l.Index) != '#' || l.sb.Len() == 0 {
		panic(l.vbSyntaxError(SyntaxError))
	}

	dateStr := strings.TrimSpace(l.sb.String())
	date, err := GetDate(dateStr)
	if err != nil {
		panic(l.vbSyntaxError(SyntaxError))
	}

	l.Index++

	return &DateLiteralToken{
		LiteralToken: LiteralToken{
			BaseToken: BaseToken{
				Start:      start,
				End:        l.Index,
				LineNumber: l.CurrentLine,
				LineStart:  l.CurrentLineStart,
			},
		},
		Value: date,
	}
}

func (l *Lexer) nextIdentifier() Token {
	start := l.Index
	id := l.getIdentifierName()

	var result Token

	switch {
	case CIEquals(id, "true"):
		result = &TrueLiteralToken{
			LiteralToken: LiteralToken{
				BaseToken: BaseToken{
					Start:      start,
					End:        l.Index,
					LineNumber: l.CurrentLine,
					LineStart:  l.CurrentLineStart,
				},
			},
		}
	case CIEquals(id, "null"):
		result = &NullLiteralToken{
			LiteralToken: LiteralToken{
				BaseToken: BaseToken{
					Start:      start,
					End:        l.Index,
					LineNumber: l.CurrentLine,
					LineStart:  l.CurrentLineStart,
				},
			},
		}
	case CIEquals(id, "false"):
		result = &FalseLiteralToken{
			LiteralToken: LiteralToken{
				BaseToken: BaseToken{
					Start:      start,
					End:        l.Index,
					LineNumber: l.CurrentLine,
					LineStart:  l.CurrentLineStart,
				},
			},
		}
	case CIEquals(id, "empty"):
		result = &EmptyLiteralToken{
			LiteralToken: LiteralToken{
				BaseToken: BaseToken{
					Start:      start,
					End:        l.Index,
					LineNumber: l.CurrentLine,
					LineStart:  l.CurrentLineStart,
				},
			},
		}
	case CIEquals(id, "nothing"):
		result = &NothingLiteralToken{
			LiteralToken: LiteralToken{
				BaseToken: BaseToken{
					Start:      start,
					End:        l.Index,
					LineNumber: l.CurrentLine,
					LineStart:  l.CurrentLineStart,
				},
			},
		}
	case IsKeyword(id):
		kw, _ := GetKeyword(id)
		result = &KeywordToken{
			BaseToken: BaseToken{
				Start:      start,
				End:        l.Index,
				LineNumber: l.CurrentLine,
				LineStart:  l.CurrentLineStart,
			},
			Keyword: kw,
			Name:    id,
		}
	case IsKeywordAsIdentifier(id):
		kw, _ := GetKeywordAsIdentifier(id)
		result = &KeywordOrIdentifierToken{
			BaseToken: BaseToken{
				Start:      start,
				End:        l.Index,
				LineNumber: l.CurrentLine,
				LineStart:  l.CurrentLineStart,
			},
			Keyword: kw,
			Name:    id,
		}
	default:
		result = &IdentifierToken{
			BaseToken: BaseToken{
				Start:      start,
				End:        l.Index,
				LineNumber: l.CurrentLine,
				LineStart:  l.CurrentLineStart,
			},
			Name: id,
		}
	}

	return result
}

func (l *Lexer) getIdentifierName() string {
	start := l.Index
	for !l.isEof() {
		c := l.getChar(l.Index)
		if IsIdentifier(c) {
			l.Index++
		} else {
			break
		}
	}
	return l.sliceString(start, l.Index)
}

func (l *Lexer) nextStringLiteral() Token {
	start := l.Index
	l.Index++ // skip opening quote

	// Fast path: ASCII-only source with no escaped quote sequences ("").
	// Scan the source bytes directly to find the closing quote without any
	// heap allocations. This is the common case for most ASP string literals.
	if l.asciiOnly {
		contentStart := l.Index
		for l.Index < l.Length {
			b := l.Code[l.Index]
			if b == '"' {
				if l.Index+1 < l.Length && l.Code[l.Index+1] == '"' {
					// Escaped quote ("") found — fall back to builder slow path.
					l.Index = start + 1
					goto slowPath
				}
				// Plain closing quote: zero-copy window into original source.
				value := l.Code[contentStart:l.Index]
				l.Index++ // advance past closing quote
				return &StringLiteralToken{
					LiteralToken: LiteralToken{
						BaseToken: BaseToken{
							Start:      start,
							End:        l.Index - 1,
							LineNumber: l.CurrentLine,
							LineStart:  l.CurrentLineStart,
						},
					},
					Value: value,
				}
			}
			if b == '\r' || b == '\n' {
				// Newline inside string — fall through to slow path (error).
				l.Index = start + 1
				goto slowPath
			}
			l.Index++
		}
		// EOF without closing quote — fall through to slow path (error).
		l.Index = start + 1
	}

slowPath:
	// Slow path: builder-based scanning for non-ASCII sources or strings
	// that contain escaped quote sequences ("").
	l.sb.Reset()
	err := true

	for !l.isEof() {
		c := l.getChar(l.Index)
		if c == '"' {
			c = l.getChar(l.Index + 1)
			l.Index++
			if c == '"' {
				l.Index++
				l.sb.WriteRune(c)
			} else {
				err = false
				break
			}
		} else if IsNewLine(c) {
			break
		} else {
			l.sb.WriteRune(c)
			l.Index++
		}
	}

	if err {
		panic(l.vbSyntaxError(UnterminatedStringConstant))
	}

	return &StringLiteralToken{
		LiteralToken: LiteralToken{
			BaseToken: BaseToken{
				Start:      start,
				End:        l.Index - 1,
				LineNumber: l.CurrentLine,
				LineStart:  l.CurrentLineStart,
			},
		},
		Value: l.sb.String(),
	}
}

func (l *Lexer) nextNumericLiteral() Token {
	start := l.Index
	c := l.getChar(l.Index)
	next := l.getChar(l.Index + 1)

	var decStr string
	var fstr strings.Builder

	if c != '.' {
		if c == '&' {
			// Skip optional whitespace between & and hex/oct prefix (Classic ASP compat: "& h22").
			nextNSIdx := l.Index + 1
			for l.getChar(nextNSIdx) == ' ' || l.getChar(nextNSIdx) == '\t' {
				nextNSIdx++
			}
			nextNS := l.getChar(nextNSIdx)
			if CharEquals(nextNS, 'h') {
				return l.nextHexIntLiteral()
			} else if CharEquals(nextNS, 'o') {
				return l.nextOctIntLiteralPrefix()
			} else if IsOctDigit(next) {
				return l.nextOctIntLiteral()
			} else {
				panic(l.vbSyntaxError(SyntaxError))
			}
		} else {
			decStr = l.getDecStr()
			if IsIdentifierStart(l.getChar(l.Index)) {
				panic(l.vbSyntaxError(ExpectedEndOfStatement))
			}
		}
	}

	c = l.getChar(l.Index)
	if c == '.' {
		l.Index++
		fstr.WriteRune('.')
		fstr.WriteString(l.getDecStr())
		c = l.getChar(l.Index)
	}

	if CharEquals(c, 'e') || CharEquals(c, 'E') {
		fstr.WriteRune('e')
		c = l.getChar(l.Index + 1)
		l.Index++
		if c == '+' || c == '-' {
			l.Index++
			fstr.WriteRune(c)
		}

		c = l.getChar(l.Index)
		if IsDecDigit(c) {
			fstr.WriteString(l.getDecStr())
		} else {
			panic(l.vbSyntaxError(InvalidNumber))
		}
	}

	c = l.getChar(l.Index)
	if IsIdentifierStart(c) {
		panic(l.vbSyntaxError(ExpectedEndOfStatement))
	}

	floatStr := fstr.String()
	if floatStr != "" && decStr != "" {
		floatStr = decStr + floatStr
	}

	if floatStr != "" {
		val, err := strconv.ParseFloat(floatStr, 64)
		if err != nil {
			panic(l.vbSyntaxError(InvalidNumber))
		}

		return &FloatLiteralToken{
			LiteralToken: LiteralToken{
				BaseToken: BaseToken{
					Start:      start,
					End:        l.Index,
					LineNumber: l.CurrentLine,
					LineStart:  l.CurrentLineStart,
				},
			},
			Value: val,
		}
	}

	result := l.parseInteger(decStr, 10)
	result.SetStart(start)
	return result
}

func (l *Lexer) getDecStr() string {
	return l.getStrByPredicate(IsDecDigit)
}

func (l *Lexer) getOctStr() string {
	return l.getStrByPredicate(IsOctDigit)
}

func (l *Lexer) getHexStr() string {
	return l.getStrByPredicate(IsHexDigit)
}

func (l *Lexer) getStrByPredicate(predicate func(rune) bool) string {
	start := l.Index
	c := l.getChar(l.Index)
	for predicate(c) {
		l.Index++
		c = l.getChar(l.Index)
	}
	return l.sliceString(start, l.Index)
}

func (l *Lexer) parseInteger(str string, base int) Token {
	val, err := strconv.ParseInt(str, base, 64)
	if err != nil {
		if base == 8 || base == 16 {
			panic(l.vbSyntaxError(SyntaxError))
		}

		floatVal, err := strconv.ParseFloat(str, 64)
		if err != nil {
			panic(l.vbSyntaxError(InvalidNumber))
		}

		return &FloatLiteralToken{
			LiteralToken: LiteralToken{
				BaseToken: BaseToken{
					End:        l.Index,
					LineNumber: l.CurrentLine,
					LineStart:  l.CurrentLineStart,
				},
			},
			Value: floatVal,
		}
	}

	var result Token
	switch base {
	case 8:
		result = &OctIntegerLiteralToken{
			DecIntegerLiteralToken: DecIntegerLiteralToken{
				LiteralToken: LiteralToken{
					BaseToken: BaseToken{
						End:        l.Index,
						LineNumber: l.CurrentLine,
						LineStart:  l.CurrentLineStart,
					},
				},
				Value: val,
			},
		}
	case 10:
		result = &DecIntegerLiteralToken{
			LiteralToken: LiteralToken{
				BaseToken: BaseToken{
					End:        l.Index,
					LineNumber: l.CurrentLine,
					LineStart:  l.CurrentLineStart,
				},
			},
			Value: val,
		}
	case 16:
		result = &HexIntegerLiteralToken{
			DecIntegerLiteralToken: DecIntegerLiteralToken{
				LiteralToken: LiteralToken{
					BaseToken: BaseToken{
						End:        l.Index,
						LineNumber: l.CurrentLine,
						LineStart:  l.CurrentLineStart,
					},
				},
				Value: val,
			},
		}
	}

	return result
}

func (l *Lexer) nextOctIntLiteralPrefix() Token {
	start := l.Index
	l.Index += 2 // skip '&o'

	str := l.getOctStr()
	c := l.getChar(l.Index)

	if IsDecDigit(c) && !IsOctDigit(c) {
		panic(l.vbSyntaxError(SyntaxError))
	}

	if IsIdentifierStart(c) {
		panic(l.vbSyntaxError(ExpectedEndOfStatement))
	}

	result := l.parseInteger(str, 8)
	result.SetStart(start)

	return result
}

func (l *Lexer) nextOctIntLiteral() Token {
	start := l.Index
	l.Index++ // skip '&'

	str := l.getOctStr()
	c := l.getChar(l.Index)

	if IsDecDigit(c) && !IsOctDigit(c) {
		panic(l.vbSyntaxError(SyntaxError))
	}

	if IsIdentifierStart(c) {
		panic(l.vbSyntaxError(ExpectedEndOfStatement))
	}

	result := l.parseInteger(str, 8)
	result.SetStart(start)

	return result
}

func (l *Lexer) nextHexIntLiteral() Token {
	start := l.Index
	l.Index++ // skip '&'
	// Skip optional whitespace between & and h (Classic ASP allows "& h22").
	for l.getChar(l.Index) == ' ' || l.getChar(l.Index) == '\t' {
		l.Index++
	}
	l.Index++ // skip 'h'

	str := l.getHexStr()
	c := l.getChar(l.Index)

	if IsIdentifierStart(c) {
		panic(l.vbSyntaxError(ExpectedEndOfStatement))
	}

	result := l.parseInteger(str, 16)
	result.SetStart(start)

	return result
}

func (l *Lexer) nextComment() Token {
	for !l.isEof() {
		c := l.getChar(l.Index)
		if c == '\'' {
			inlineComment := l.hasCodeBeforeCurrentToken()
			l.Index++
			return l.nextCommentBody(1, false, inlineComment)
		} else if CharEquals(c, 'r') {
			c2 := l.getChar(l.Index + 1)
			c3 := l.getChar(l.Index + 2)
			c4 := l.getChar(l.Index + 3)
			if CharEquals(c2, 'e') && CharEquals(c3, 'm') && IsWhiteSpace(c4) {
				inlineComment := l.hasCodeBeforeCurrentToken()
				l.Index += 3
				return l.nextCommentBody(3, true, inlineComment)
			}
			break
		} else {
			break
		}
	}
	return nil
}

func (l *Lexer) nextCommentBody(offset int, isRem bool, inlineComment bool) Token {
	start := l.Index - offset
	l.sb.Reset()

	for !l.isEof() {
		c := l.getChar(l.Index)
		if inlineComment && l.Mode == ModeASP && l.InASPBlock && l.BlockType == BlockTypePercent {
			next := l.getChar(l.Index + 1)
			if c == '%' && next == '>' {
				break
			}
		}
		if IsNewLine(c) {
			break
		}
		l.sb.WriteRune(c)
		l.Index++
	}

	return &CommentToken{
		BaseToken: BaseToken{
			Start:      start,
			End:        l.Index,
			LineNumber: l.CurrentLine,
			LineStart:  l.CurrentLineStart,
		},
		Comment: l.sb.String(),
		IsRem:   isRem,
	}
}

// hasCodeBeforeCurrentToken reports whether there is non-whitespace content
// earlier on the same line before the current token start.
func (l *Lexer) hasCodeBeforeCurrentToken() bool {
	for i := l.CurrentLineStart; i < l.Index; i++ {
		ch := l.getChar(i)
		if ch != ' ' && ch != '\t' {
			return true
		}
	}
	return false
}

func (l *Lexer) isScriptServerStart() (int, bool) {
	if l.getChar(l.Index) != '<' {
		return 0, false
	}
	s := l.sliceString(l.Index, l.Index+7)
	if !strings.EqualFold(s, "<script") {
		return 0, false
	}

	// Find the end of the opening tag >
	for i := l.Index + 7; i < l.Length; i++ {
		c := l.getChar(i)
		if c == '>' {
			attr := l.sliceString(l.Index+7, i)
			attrLower := strings.ToLower(attr)

			// Must have runat="server"
			if !strings.Contains(attrLower, "runat=\"server\"") &&
				!strings.Contains(attrLower, "runat=server") &&
				!strings.Contains(attrLower, "runat='server'") {
				return 0, false
			}

			// Optional language="vbscript" check (can be skipped as we assume VBScript in this engine)
			return i - l.Index + 1, true
		}
	}
	return 0, false
}

func (l *Lexer) isObjectStart() (int, bool, map[string]string) {
	if l.getChar(l.Index) != '<' {
		return 0, false, nil
	}
	s := l.sliceString(l.Index, l.Index+7)
	if !strings.EqualFold(s, "<object") {
		return 0, false, nil
	}

	// Find the end of the tag >
	for i := l.Index + 7; i < l.Length; i++ {
		if l.getChar(i) == '>' {
			attrStr := l.sliceString(l.Index+7, i)
			attrLower := strings.ToLower(attrStr)

			if !strings.Contains(attrLower, "runat=\"server\"") &&
				!strings.Contains(attrLower, "runat=server") &&
				!strings.Contains(attrLower, "runat='server'") {
				return 0, false, nil
			}

			// Parse attributes simply
			attrs := make(map[string]string)
			// Find scope, id, progid, classid
			for _, key := range []string{"scope", "id", "progid", "classid"} {
				if idx := strings.Index(attrLower, key); idx != -1 {
					// find next = then quotes
					sub := attrStr[idx+len(key):]
					if eqIdx := strings.Index(sub, "="); eqIdx != -1 {
						valSub := sub[eqIdx+1:]
						valSub = strings.TrimSpace(valSub)
						if len(valSub) > 0 {
							quote := valSub[0]
							if quote == '"' || quote == '\'' {
								if endQuote := strings.Index(valSub[1:], string(quote)); endQuote != -1 {
									attrs[key] = valSub[1 : 1+endQuote]
								}
							} else {
								// unquoted?
								endSpace := strings.IndexAny(valSub, " >")
								if endSpace == -1 {
									attrs[key] = valSub
								} else {
									attrs[key] = valSub[:endSpace]
								}
							}
						}
					}
				}
			}

			// Now find the closing </object>
			for j := i + 1; j < l.Length; j++ {
				if l.getChar(j) == '<' && l.getChar(j+1) == '/' {
					closeTag := l.sliceString(j, j+9)
					if strings.EqualFold(closeTag, "</object>") {
						return j - l.Index + 9, true, attrs
					}
				}
			}

			return i - l.Index + 1, true, attrs
		}
	}
	return 0, false, nil
}

func (l *Lexer) isScriptEnd() (int, bool) {
	if l.getChar(l.Index) != '<' || l.getChar(l.Index+1) != '/' {
		return 0, false
	}
	s := l.sliceString(l.Index, l.Index+9)
	if strings.EqualFold(s, "</script>") {
		return 9, true
	}
	return 0, false
}

func (l *Lexer) isIncludeStart() (int, bool, string, bool) {
	if l.getChar(l.Index) != '<' || l.getChar(l.Index+1) != '!' || l.getChar(l.Index+2) != '-' || l.getChar(l.Index+3) != '-' || l.getChar(l.Index+4) != '#' {
		return 0, false, "", false
	}
	s := l.sliceString(l.Index, l.Index+13)
	if !strings.EqualFold(s, "<!--#include ") {
		return 0, false, "", false
	}

	// Find the end -->
	for i := l.Index + 13; i < l.Length; i++ {
		if l.getChar(i) == '-' && l.getChar(i+1) == '-' && l.getChar(i+2) == '>' {
			content := l.sliceString(l.Index+13, i)
			contentLower := strings.ToLower(content)
			virtual := strings.Contains(contentLower, "virtual")

			// Find path between quotes
			var path string
			if q1 := strings.Index(content, "\""); q1 != -1 {
				if q2 := strings.Index(content[q1+1:], "\""); q2 != -1 {
					path = content[q1+1 : q1+1+q2]
				}
			} else if q1 := strings.Index(content, "'"); q1 != -1 {
				if q2 := strings.Index(content[q1+1:], "'"); q2 != -1 {
					path = content[q1+1 : q1+1+q2]
				}
			}

			return i - l.Index + 3, true, path, virtual
		}
	}
	return 0, false, "", false
}

// consumeASPPostBlockLineBreak applies Classic ASP newline suppression after one
// code block terminator (%> or </script>). IIS suppresses one immediate line break,
// including cases where only spaces/tabs appear before that line break.
func (l *Lexer) consumeASPPostBlockLineBreak() {
	if !l.skipHTMLLeadingNL {
		return
	}

	// Fast path for code-only formatting between consecutive server-side blocks:
	// collapse all horizontal/newline whitespace when the next non-whitespace
	// token starts another server delimiter.
	origIndex := l.Index
	tmpIndex := l.Index
	tmpLine := l.CurrentLine
	tmpLineStart := l.CurrentLineStart
	for {
		ch := l.getChar(tmpIndex)
		if ch == ' ' || ch == '\t' {
			tmpIndex++
			continue
		}
		if ch == '\r' {
			tmpIndex++
			if l.getChar(tmpIndex) == '\n' {
				tmpIndex++
			}
			tmpLine++
			tmpLineStart = tmpIndex
			continue
		}
		if ch == '\n' {
			tmpIndex++
			tmpLine++
			tmpLineStart = tmpIndex
			continue
		}
		break
	}
	if tmpIndex > origIndex && l.startsServerDelimiterAt(tmpIndex) {
		l.Index = tmpIndex
		l.CurrentLine = tmpLine
		l.CurrentLineStart = tmpLineStart
		l.skipHTMLLeadingNL = false
		return
	}

	i := l.Index
	for {
		ch := l.getChar(i)
		if ch == ' ' || ch == '\t' {
			i++
			continue
		}
		break
	}

	ch := l.getChar(i)
	if ch == '\r' {
		i++
		if l.getChar(i) == '\n' {
			i++
		}
		l.Index = i
		l.CurrentLine++
		l.CurrentLineStart = l.Index
	} else if ch == '\n' {
		i++
		l.Index = i
		l.CurrentLine++
		l.CurrentLineStart = l.Index
	}

	l.skipHTMLLeadingNL = false
}

// startsServerDelimiterAt reports whether a given source position begins a
// server-side ASP delimiter recognized by this lexer in ASP mode.
func (l *Lexer) startsServerDelimiterAt(pos int) bool {
	if l.getChar(pos) == '<' && l.getChar(pos+1) == '%' {
		return true
	}
	if l.getChar(pos) == 0 {
		return true
	}
	prev := l.Index
	l.Index = pos
	_, isScript := l.isScriptServerStart()
	if isScript {
		l.Index = prev
		return true
	}
	_, isInclude, _, _ := l.isIncludeStart()
	if isInclude {
		l.Index = prev
		return true
	}
	_, isObject, _ := l.isObjectStart()
	l.Index = prev
	return isObject
}

// shouldSuppressFormattingHTMLBeforeServerBlock reports whether one HTML segment
// before a server-side delimiter is only formatting whitespace with line breaks.
// This avoids emitting blank-line noise from code-only formatting around ASP blocks.
func (l *Lexer) shouldSuppressFormattingHTMLBeforeServerBlock(start int, end int) bool {
	if end <= start {
		return false
	}
	hasNewline := false
	for i := start; i < end; i++ {
		ch := l.getChar(i)
		if ch == '\r' || ch == '\n' {
			hasNewline = true
			continue
		}
		if ch == ' ' || ch == '\t' {
			continue
		}
		return false
	}
	return hasNewline
}

func (l *Lexer) nextHTML() Token {
	l.consumeASPPostBlockLineBreak()

	start := l.Index
	line := l.CurrentLine
	lineStart := l.CurrentLineStart

	for !l.isEof() {
		c := l.getChar(l.Index)

		// Check for <%
		if c == '<' && l.getChar(l.Index+1) == '%' {
			if l.Index > start {
				if l.shouldSuppressFormattingHTMLBeforeServerBlock(start, l.Index) {
					start = l.Index
					line = l.CurrentLine
					lineStart = l.CurrentLineStart
				} else {
					return &HTMLToken{
						BaseToken: BaseToken{
							Start:      start,
							End:        l.Index,
							LineNumber: line,
							LineStart:  lineStart,
						},
						Content: l.sliceString(start, l.Index),
					}
				}
			}

			// It's the start of an ASP block
			aspStart := l.Index
			aspLine := l.CurrentLine
			aspLineStart := l.CurrentLineStart
			l.Index += 2
			l.InASPBlock = true
			l.BlockType = BlockTypePercent

			// Classic ASP accepts optional whitespace/newlines between '<%' and
			// the directive/expression marker ('@' or '=').
			probe := l.Index
			for !l.isEof() {
				ch := l.getChar(probe)
				if ch == ' ' || ch == '\t' || ch == '\r' || ch == '\n' {
					probe++
					continue
				}
				break
			}

			next := l.getChar(probe)
			if next == '=' {
				l.Index = probe + 1
				return &ASPExpressionStartToken{
					BaseToken: BaseToken{
						Start:      aspStart,
						End:        l.Index,
						LineNumber: aspLine,
						LineStart:  aspLineStart,
					},
				}
			} else if next == '@' {
				l.Index = probe + 1
				return &ASPDirectiveStartToken{
					BaseToken: BaseToken{
						Start:      aspStart,
						End:        l.Index,
						LineNumber: aspLine,
						LineStart:  aspLineStart,
					},
				}
			}
			return &ASPCodeStartToken{
				BaseToken: BaseToken{
					Start:      aspStart,
					End:        l.Index,
					LineNumber: aspLine,
					LineStart:  aspLineStart,
				},
			}
		}

		// Check for <script runat=server>
		if length, ok := l.isScriptServerStart(); ok {
			if l.Index > start {
				if l.shouldSuppressFormattingHTMLBeforeServerBlock(start, l.Index) {
					start = l.Index
					line = l.CurrentLine
					lineStart = l.CurrentLineStart
				} else {
					return &HTMLToken{
						BaseToken: BaseToken{
							Start:      start,
							End:        l.Index,
							LineNumber: line,
							LineStart:  lineStart,
						},
						Content: l.sliceString(start, l.Index),
					}
				}
			}

			aspStart := l.Index
			aspLine := l.CurrentLine
			aspLineStart := l.CurrentLineStart
			l.Index += length
			l.InASPBlock = true
			l.BlockType = BlockTypeScript
			return &ASPCodeStartToken{
				BaseToken: BaseToken{
					Start:      aspStart,
					End:        l.Index,
					LineNumber: aspLine,
					LineStart:  aspLineStart,
				},
			}
		}

		// Check for <!--#include
		if length, ok, path, virtual := l.isIncludeStart(); ok {
			if l.Index > start {
				if l.shouldSuppressFormattingHTMLBeforeServerBlock(start, l.Index) {
					start = l.Index
					line = l.CurrentLine
					lineStart = l.CurrentLineStart
				} else {
					return &HTMLToken{
						BaseToken: BaseToken{
							Start:      start,
							End:        l.Index,
							LineNumber: line,
							LineStart:  lineStart,
						},
						Content: l.sliceString(start, l.Index),
					}
				}
			}

			aspStart := l.Index
			aspLine := l.CurrentLine
			aspLineStart := l.CurrentLineStart
			l.Index += length
			l.skipHTMLLeadingNL = true
			return &ASPIncludeToken{
				BaseToken: BaseToken{
					Start:      aspStart,
					End:        l.Index,
					LineNumber: aspLine,
					LineStart:  aspLineStart,
				},
				Path:    path,
				Virtual: virtual,
			}
		}

		// Check for <object
		if length, ok, attrs := l.isObjectStart(); ok {
			if l.Index > start {
				if l.shouldSuppressFormattingHTMLBeforeServerBlock(start, l.Index) {
					start = l.Index
					line = l.CurrentLine
					lineStart = l.CurrentLineStart
				} else {
					return &HTMLToken{
						BaseToken: BaseToken{
							Start:      start,
							End:        l.Index,
							LineNumber: line,
							LineStart:  lineStart,
						},
						Content: l.sliceString(start, l.Index),
					}
				}
			}

			aspStart := l.Index
			aspLine := l.CurrentLine
			aspLineStart := l.CurrentLineStart
			l.Index += length
			return &ASPObjectToken{
				BaseToken: BaseToken{
					Start:      aspStart,
					End:        l.Index,
					LineNumber: aspLine,
					LineStart:  aspLineStart,
				},
				Scope:   attrs["scope"],
				ID:      attrs["id"],
				ProgID:  attrs["progid"],
				ClassID: attrs["classid"],
			}
		}

		if IsNewLine(c) {
			if c == '\r' && l.getChar(l.Index+1) == '\n' {
				l.Index++
			}
			l.CurrentLine++
			l.Index++
			l.CurrentLineStart = l.Index
		} else {
			l.Index++
		}
	}

	if l.Index > start {
		return &HTMLToken{
			BaseToken: BaseToken{
				Start:      start,
				End:        l.Index,
				LineNumber: line,
				LineStart:  lineStart,
			},
			Content: l.sliceString(start, l.Index),
		}
	}

	return &EOFToken{
		BaseToken: BaseToken{
			Start:      l.Index,
			End:        l.Index,
			LineNumber: l.CurrentLine,
			LineStart:  l.CurrentLineStart,
		},
	}
}

func (l *Lexer) nextLineTermination() Token {
	start := l.Index
	line := l.CurrentLine
	isColon := false

	for !l.isEof() {
		c := l.getChar(l.Index)
		if IsLineTerminator(c) {
			if c == '\r' && l.getChar(l.Index+1) == '\n' {
				l.Index++
			}

			l.Index++
			isColon = isColon || (c == ':')

			if c != ':' {
				l.CurrentLine++
				l.CurrentLineStart = l.Index
			}
		} else {
			break
		}

		l.skipWhitespaces()
	}

	var token Token
	if isColon && line == l.CurrentLine {
		token = &ColonLineTerminationToken{
			LineTerminationToken: LineTerminationToken{
				BaseToken: BaseToken{
					Start:      start,
					End:        l.Index,
					LineNumber: l.CurrentLine - 1,
					LineStart:  l.CurrentLineStart,
				},
			},
		}
	} else {
		token = &LineTerminationToken{
			BaseToken: BaseToken{
				Start:      start,
				End:        l.Index,
				LineNumber: l.CurrentLine - 1,
				LineStart:  l.CurrentLineStart,
			},
		}
	}

	return token
}

// getNextNonWhitespaceChar returns the next non-whitespace character after the current position
// This is used for detecting operators that may have whitespace between components
func (l *Lexer) getNextNonWhitespaceChar(startPos int) rune {
	pos := startPos
	for pos < l.Length {
		ch := l.getChar(pos)
		// Allow only space and tab as whitespace, not newlines
		if ch != ' ' && ch != '\t' {
			return ch
		}
		pos++
	}
	return 0
}

func (l *Lexer) nextPunctuation() Token {
	start := l.Index
	c := l.getChar(l.Index)
	next := l.getChar(l.Index + 1)
	// For compound operators, check the next non-whitespace character
	nextNonWS := l.getNextNonWhitespaceChar(l.Index + 1)

	var punctType *Punctuation

	switch c {
	case '(':
		p := PunctLParen
		punctType = &p
	case ')':
		p := PunctRParen
		punctType = &p
	case '.':
		p := PunctDot
		punctType = &p
	case ',':
		p := PunctComma
		punctType = &p
	case '+':
		p := PunctPlus
		punctType = &p
	case '-':
		p := PunctMinus
		punctType = &p
	case '/':
		p := PunctSlash
		punctType = &p
	case '\\':
		p := PunctBackslash
		punctType = &p
	case '*':
		p := PunctStar
		punctType = &p
	case '&':
		p := PunctAmp
		punctType = &p
	case '^':
		p := PunctExp
		punctType = &p
	case '=':
		if next == '<' {
			l.Index++
			p := PunctLessOrEqual
			punctType = &p
		} else if next == '>' {
			l.Index++
			p := PunctGreaterOrEqual
			punctType = &p
		} else {
			p := PunctEqual
			punctType = &p
		}
	case '<':
		if next == '=' {
			l.Index++
			p := PunctLessOrEqual
			punctType = &p
		} else if next == '>' {
			l.Index++
			p := PunctNotEqual
			punctType = &p
		} else if nextNonWS == '=' {
			// Handle < = (with spaces)
			// Count how many characters to skip
			pos := l.Index + 1
			for pos < l.Length {
				ch := l.getChar(pos)
				if ch != ' ' && ch != '\t' {
					break
				}
				pos++
			}
			if pos < l.Length && l.getChar(pos) == '=' {
				l.Index = pos // Move to the '=' position, will be incremented below
				p := PunctLessOrEqual
				punctType = &p
			} else {
				p := PunctLess
				punctType = &p
			}
		} else {
			p := PunctLess
			punctType = &p
		}
	case '>':
		if next == '=' {
			l.Index++
			p := PunctGreaterOrEqual
			punctType = &p
		} else if next == '<' {
			l.Index++
			p := PunctNotEqual
			punctType = &p
		} else if nextNonWS == '=' {
			// Handle > = (with spaces)
			// Count how many characters to skip
			pos := l.Index + 1
			for pos < l.Length {
				ch := l.getChar(pos)
				if ch != ' ' && ch != '\t' {
					break
				}
				pos++
			}
			if pos < l.Length && l.getChar(pos) == '=' {
				l.Index = pos // Move to the '=' position, will be incremented below
				p := PunctGreaterOrEqual
				punctType = &p
			} else {
				p := PunctGreater
				punctType = &p
			}
		} else {
			p := PunctGreater
			punctType = &p
		}
	}

	if punctType == nil {
		panic(l.vbSyntaxError(InvalidCharacter))
	}

	l.Index++

	return &PunctuationToken{
		BaseToken: BaseToken{
			Start:      start,
			End:        l.Index,
			LineNumber: l.CurrentLine,
			LineStart:  l.CurrentLineStart,
		},
		Type: *punctType,
	}
}

func (l *Lexer) vbSyntaxError(code VBSyntaxErrorCode) error {
	// Capture the current offending character/token and full line text
	tokenText := ""
	if l.Index < l.Length {
		r := l.getChar(l.Index)
		if r != 0 {
			tokenText = string(r)
		}
	}

	lineText := l.currentLineText()

	return NewVBSyntaxError(code, l.CurrentLine, l.LineIndex(), tokenText, lineText)
}

// VBSyntaxError represents a VBScript syntax error
type VBSyntaxError struct {
	Code           VBSyntaxErrorCode
	Line           int
	Column         int
	TokenText      string
	LineText       string
	ASPCode        int
	ASPDescription string
	Category       string
	Description    string
	File           string
	Number         int
	Source         string
}

// NewVBSyntaxError creates a new syntax error
func NewVBSyntaxError(code VBSyntaxErrorCode, line, column int, tokenText, lineText string) *VBSyntaxError {
	description := code.String()
	return &VBSyntaxError{
		Code:           code,
		Line:           line,
		Column:         normalizeVBScriptColumn(column),
		TokenText:      tokenText,
		LineText:       lineText,
		ASPCode:        int(code),
		ASPDescription: description,
		Category:       "VBScript compilation",
		Description:    description,
		File:           "",
		Number:         HRESULTFromVBScriptCode(code),
		Source:         "VBScript compilation error",
	}
}

// WithFile attaches a source file path to the syntax error.
func (e *VBSyntaxError) WithFile(file string) *VBSyntaxError {
	if e == nil {
		return nil
	}

	e.File = strings.TrimSpace(file)
	return e
}

// WithASPDescription overrides the ASP description while keeping the catalog description intact.
func (e *VBSyntaxError) WithASPDescription(description string) *VBSyntaxError {
	if e == nil {
		return nil
	}

	trimmed := strings.TrimSpace(description)
	if trimmed != "" {
		e.ASPDescription = trimmed
	}
	return e
}

// Error implements the error interface
func (e *VBSyntaxError) Error() string {
	if e == nil {
		return ""
	}

	var builder strings.Builder
	builder.Grow(256)
	builder.WriteString(e.Source)
	if e.Code != 0 {
		builder.WriteString(" '")
		builder.WriteString(HRESULTHexFromVBScriptCode(e.Code))
		builder.WriteString("'")
	}

	if strings.TrimSpace(e.Description) != "" {
		builder.WriteString("\n")
		builder.WriteString(e.Description)
	}

	builder.WriteString("\nCategory: ")
	builder.WriteString(e.Category)
	builder.WriteString("\nColumn: ")
	builder.WriteString(strconv.Itoa(e.Column))
	builder.WriteString("\nDescription: ")
	builder.WriteString(e.Description)
	builder.WriteString("\nFile: ")
	builder.WriteString(e.File)
	builder.WriteString("\nLine: ")
	builder.WriteString(strconv.Itoa(e.Line))
	builder.WriteString("\nNumber: ")
	builder.WriteString(strconv.Itoa(e.Number))
	builder.WriteString("\nSource: ")
	builder.WriteString(e.Source)

	if e.LineText != "" {
		builder.WriteString("\n\n")
		builder.WriteString(e.LineText)
	}

	return builder.String()
}

// normalizeVBScriptColumn converts the internal zero-based column into a VBScript-compatible one-based column.
func normalizeVBScriptColumn(column int) int {
	if column < 0 {
		return 0
	}

	return column + 1
}

// currentLineText returns the full text of the current line
func (l *Lexer) currentLineText() string {
	if len(l.Code) == 0 {
		return ""
	}
	// Find start and end of current line using rune indices
	start := l.CurrentLineStart
	if start < 0 {
		start = 0
	}
	// Scan forward until newline or EOF
	end := start
	for end < l.Length {
		ch := l.getChar(end)
		if ch == '\n' || ch == '\r' || ch == 0 {
			break
		}
		end++
	}
	if start >= 0 && start < l.Length && end <= l.Length && end >= start {
		return l.sliceString(start, end)
	}
	return ""
}
