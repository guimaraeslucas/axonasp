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
	"strings"
	"time"
)

var (
	// keywords maps lowercase keyword strings to their keyword enum values
	keywords = map[string]Keyword{
		"and":      KeywordAnd,
		"byref":    KeywordByRef,
		"byval":    KeywordByVal,
		"call":     KeywordCall,
		"case":     KeywordCase,
		"class":    KeywordClass,
		"const":    KeywordConst,
		"dim":      KeywordDim,
		"do":       KeywordDo,
		"each":     KeywordEach,
		"else":     KeywordElse,
		"elseif":   KeywordElseIf,
		"end":      KeywordEnd,
		"eqv":      KeywordEqv,
		"exit":     KeywordExit,
		"for":      KeywordFor,
		"function": KeywordFunction,
		"get":      KeywordGet,
		"goto":     KeywordGoto,
		"if":       KeywordIf,
		"imp":      KeywordImp,
		"in":       KeywordIn,
		"is":       KeywordIs,
		"let":      KeywordLet,
		"loop":     KeywordLoop,
		"mod":      KeywordMod,
		"new":      KeywordNew,
		"next":     KeywordNext,
		"not":      KeywordNot,
		"on":       KeywordOn,
		"option":   KeywordOption,
		"or":       KeywordOr,
		"preserve": KeywordPreserve,
		"private":  KeywordPrivate,
		"public":   KeywordPublic,
		"redim":    KeywordReDim,
		"resume":   KeywordResume,
		"select":   KeywordSelect,
		"set":      KeywordSet,
		"sub":      KeywordSub,
		"then":     KeywordThen,
		"to":       KeywordTo,
		"until":    KeywordUntil,
		"wend":     KeywordWEnd,
		"while":    KeywordWhile,
		"with":     KeywordWith,
		"xor":      KeywordXor,
	}

	// keywordAsIdentifiers maps keyword-like strings that can also be used as identifiers
	keywordAsIdentifiers = map[string]Keyword{
		"binary":   KeywordBinary,
		"compare":  KeywordCompare,
		"default":  KeywordDefault,
		"erase":    KeywordErase,
		"error":    KeywordError,
		"explicit": KeywordExplicit,
		"base":     KeywordBase,
		"property": KeywordProperty,
		"step":     KeywordStep,
		"text":     KeywordText,
	}
)

// IsKeyword checks if a string is a reserved keyword
func IsKeyword(s string) bool {
	_, ok := keywords[strings.ToLower(s)]
	return ok
}

// GetKeyword returns the keyword enum value for a string
func GetKeyword(s string) (Keyword, bool) {
	kw, ok := keywords[strings.ToLower(s)]
	return kw, ok
}

// IsKeywordAsIdentifier checks if a string is a keyword that can be used as an identifier
func IsKeywordAsIdentifier(s string) bool {
	_, ok := keywordAsIdentifiers[strings.ToLower(s)]
	return ok
}

// GetKeywordAsIdentifier returns the keyword enum value for a keyword-as-identifier string
func GetKeywordAsIdentifier(s string) (Keyword, bool) {
	kw, ok := keywordAsIdentifiers[strings.ToLower(s)]
	return kw, ok
}

// GetDate parses a date string and returns a time.Time
// Note: This is a basic implementation. VBScript date parsing may be more complex.
func GetDate(s string) (time.Time, error) {
	// Try common date formats
	// Use '1' and '2' for month/day to allow single digits (Go matches "01" with "1" pattern too)
	formats := []string{
		"1/2/2006",
		"2006-1-2",
		"2006-01-02",
		"01/02/2006",
		"1/2/2006 3:04:05 PM",
		"1/2/2006 15:04:05",
		"2006-1-2 15:04:05",
		"2006-01-02 15:04:05",
		"3:04:05 PM",
		"15:04:05",
		time.RFC3339,
		time.RFC3339Nano,
	}

	for _, format := range formats {
		if t, err := time.Parse(format, s); err == nil {
			return t, nil
		}
	}

	// Fall back to Go's time.Parse which handles many common formats
	return time.Parse(time.RFC3339, s)
}
