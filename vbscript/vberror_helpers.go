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
 */
package vbscript

import "fmt"

const vbscriptHRESULTBase uint32 = 0x800A0000

// HRESULTFromVBScriptCode converts a VBScript catalog code to a classic HRESULT value.
func HRESULTFromVBScriptCode(code VBSyntaxErrorCode) int {
	return int(int32(vbscriptHRESULTBase | uint32(uint16(code))))
}

// HRESULTHexFromVBScriptCode formats a VBScript HRESULT using the classic 8-digit uppercase representation.
func HRESULTHexFromVBScriptCode(code VBSyntaxErrorCode) string {
	return fmt.Sprintf("%08X", uint32(vbscriptHRESULTBase|uint32(uint16(code))))
}
