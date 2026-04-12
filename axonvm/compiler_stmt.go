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
)

// In axonvm/compiler_expr.go we have expression compilation.
// We'll add statement and control flow compilation here.

// emitJump emits a jump instruction with a dummy 0 offset and returns the offset position
// so it can be patched later.
func (c *Compiler) emitJump(op OpCode) int {
	c.emit(op, 0)
	return len(c.bytecode) - 4 // The 32-bit jump target starts 4 bytes from the end
}

// patchJump updates the jump offset at the given index to jump to the current bytecode length.
func (c *Compiler) patchJump(offsetIndex int) {
	c.patchJumpTo(offsetIndex, len(c.bytecode))
}

// patchJumpTo updates one jump offset to a specific absolute bytecode index.
func (c *Compiler) patchJumpTo(offsetIndex int, jumpTarget int) {
	if offsetIndex <= 0 || offsetIndex > len(c.bytecode) {
		panic("Invalid jump patch offset")
	}
	if usesWideJumpOperand(OpCode(c.bytecode[offsetIndex-1])) {
		binary.BigEndian.PutUint32(c.bytecode[offsetIndex:], uint32(jumpTarget))
		return
	}
	binary.BigEndian.PutUint16(c.bytecode[offsetIndex:], uint16(jumpTarget))
}

func (c *Compiler) emitLoop(loopStart int) {
	c.emit(OpJump, loopStart)
}
