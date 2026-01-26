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

// IsLineTerminator checks if a character is a line terminator
func IsLineTerminator(c rune) bool {
	return c == '\n' || c == '\r' || c == ':'
}

// IsNewLine checks if a character is a new line character
func IsNewLine(c rune) bool {
	return c == '\n' || c == '\r'
}

// IsDecDigit checks if a character is a decimal digit (0-9)
func IsDecDigit(c rune) bool {
	return c >= '0' && c <= '9'
}

// IsHexDigit checks if a character is a hexadecimal digit (0-9, A-F, a-f)
func IsHexDigit(c rune) bool {
	return IsDecDigit(c) ||
		(c >= 'a' && c <= 'f') ||
		(c >= 'A' && c <= 'F')
}

// IsOctDigit checks if a character is an octal digit (0-7)
func IsOctDigit(c rune) bool {
	return c >= '0' && c <= '7'
}

// IsIdentifierStart checks if a character can start an identifier
func IsIdentifierStart(c rune) bool {
	return (c >= 'A' && c <= 'Z') ||
		(c >= 'a' && c <= 'z')
}

// IsIdentifier checks if a character is valid in an identifier
func IsIdentifier(c rune) bool {
	return IsIdentifierStart(c) ||
		IsDecDigit(c) ||
		c == '_'
}

// IsWhiteSpace checks if a character is whitespace
func IsWhiteSpace(c rune) bool {
	return c == 0x20 || c == 0x09 ||
		c == 0x0B || c == 0x0C
}

// IsExtendedIdentifier checks if a character is valid in an extended identifier
func IsExtendedIdentifier(c rune) bool {
	return !IsNewLine(c) && c != ']' && c >= 0 && c <= 0xff
}

// CharEquals compares two characters case-insensitively
func CharEquals(a, b rune) bool {
	return runeToUpper(a) == runeToUpper(b)
}

// runeToUpper converts a rune to uppercase
func runeToUpper(c rune) rune {
	if c >= 'a' && c <= 'z' {
		return c - 32
	}
	return c
}
