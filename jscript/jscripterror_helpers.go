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
 */
package jscript

import "fmt"

const jscriptHRESULTBase uint32 = 0x800A0000

// HRESULTFromJScriptCode converts a JScript catalog code to a classic HRESULT value.
func HRESULTFromJScriptCode(code JSSyntaxErrorCode) int {
	return int(int32(jscriptHRESULTBase | uint32(uint16(code))))
}

// HRESULTHexFromJScriptCode formats a JScript HRESULT using the classic 8-digit uppercase representation.
func HRESULTHexFromJScriptCode(code JSSyntaxErrorCode) string {
	return fmt.Sprintf("%08X", uint32(jscriptHRESULTBase|uint32(uint16(code))))
}
