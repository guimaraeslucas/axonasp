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
)

// GetChar returns the character at a given position in a string, or rune(0) if out of bounds
func GetChar(str string, pos int) rune {
	if pos < 0 || pos >= len(str) {
		return rune(0)
	}
	return []rune(str)[pos]
}

// GetCharCode returns the numeric code of a character at a given position
func GetCharCode(str string, pos int) int {
	if pos < 0 || pos >= len(str) {
		return 0
	}
	return int([]rune(str)[pos])
}

// Slice returns a substring from start to end position
// Supports negative indices like JavaScript slice
// end is exclusive
func Slice(str string, start, end int) string {
	runes := []rune(str)
	length := len(runes)

	// Convert negative indices
	from := start
	if from < 0 {
		from = length + from
		if from < 0 {
			from = 0
		}
	}
	if from > length {
		from = length
	}

	to := end
	if to < 0 {
		to = length + to
		if to < 0 {
			to = 0
		}
	}
	if to > length {
		to = length
	}

	// Ensure from <= to
	if from > to {
		return ""
	}

	return string(runes[from:to])
}

// CIEquals performs case-insensitive string comparison
func CIEquals(s1, s2 string) bool {
	return strings.EqualFold(s1, s2)
}

// StringToLower converts a string to lowercase (using English locale)
func StringToLower(s string) string {
	return strings.ToLower(s)
}

// StringToUpper converts a string to uppercase (using English locale)
func StringToUpper(s string) string {
	return strings.ToUpper(s)
}
