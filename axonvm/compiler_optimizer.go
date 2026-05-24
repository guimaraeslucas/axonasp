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
	"math"
	"math/bits"
	"strconv"
)

// optimizePeephole performs in-place bytecode peephole optimization.
// It repeats single-pass scans until no further constant folding is possible,
// allowing chained binary operations (e.g. 1+2+3) to fully collapse.
// All changes are made in-place on c.bytecode; redundant bytes are replaced
// with OpNop so every absolute jump offset remains valid.
func (c *Compiler) optimizePeephole() {
	for {
		folded := c.optimizePeepholePass()
		propagated := c.optimizeLocalCopyPropagationPass()
		intOptimized := c.optimizeIntegerArithmeticPass()
		deadCode := c.optimizeDeadConditionalJumpPass()
		fusedBranch := c.optimizeFusedBranchPass()
		loadBranch := c.optimizeFusedLoadBranchPass()
		inPlaceMath := c.optimizeInPlaceMathPass()
		constPooling := c.optimizeConstantPoolingPass()
		if !folded && !propagated && !intOptimized && !deadCode && !fusedBranch && !loadBranch && !inPlaceMath && !constPooling {
			break
		}
	}
}

// optimizeFusedBranchPass merges comparison opcodes followed by OpJumpIfFalse
// into single fused branch super-instructions.
func (c *Compiler) optimizeFusedBranchPass() bool {
	if len(c.bytecode) < 6 {
		return false
	}
	targets := collectJumpTargets(c.bytecode)
	changed := false

	for i := 0; i < len(c.bytecode); {
		op := OpCode(c.bytecode[i])
		if !isFusedBranchCandidateOp(op) {
			i++
			continue
		}

		// Skip any OpNop padding to find the jump instruction.
		j := i + 1
		for j < len(c.bytecode) && OpCode(c.bytecode[j]) == OpNop {
			j++
		}
		// Fused branch candidates are 1-byte opcodes. Jumps are 5 bytes.
		if j+5 > len(c.bytecode) {
			i++
			continue
		}

		jumpOp := OpCode(c.bytecode[j])
		if !isFusedBranchJumpOp(jumpOp) {
			i++
			continue
		}

		// Safety: no jump target may land on the padding bytes between comparison and jump.
		if hasTargetInRange(targets, i+1, j) {
			i++
			continue
		}

		fusedOp := getFusedBranchOp(op, jumpOp)
		if fusedOp == OpHalt { // Sentinel for not foldable
			i++
			continue
		}

		// Read the 4-byte absolute target from the jump opcode.
		target := binary.BigEndian.Uint32(c.bytecode[j+1 : j+5])

		// Replace the comparison opcode with the fused branch opcode and its target.
		c.bytecode[i] = byte(fusedOp)
		binary.BigEndian.PutUint32(c.bytecode[i+1:i+5], target)

		// fill every byte from i+5 through j+4 (inclusive) with OpNop.
		for p := i + 5; p <= j+4; p++ {
			c.bytecode[p] = byte(OpNop)
		}

		changed = true
		i = j + 5
	}
	return changed
}

func isFusedBranchCandidateOp(op OpCode) bool {
	switch op {
	case OpEq, OpNeq, OpLt, OpGt, OpIsRef,
		OpJSLooseEqual, OpJSLooseNotEqual, OpJSStrictEq, OpJSStrictNeq, OpJSLess, OpJSGreater:
		return true
	}
	return false
}

func isFusedBranchJumpOp(op OpCode) bool {
	return op == OpJumpIfFalse || op == OpJSJumpIfFalse
}

func getFusedBranchOp(op OpCode, jumpOp OpCode) OpCode {
	if jumpOp == OpJumpIfFalse {
		switch op {
		case OpEq:
			return OpJumpIfNotEq
		case OpNeq:
			return OpJumpIfEq
		case OpLt:
			return OpJumpIfNotLt
		case OpGt:
			return OpJumpIfLte
		case OpIsRef:
			return OpJumpIfNotIs
		}
	} else if jumpOp == OpJSJumpIfFalse {
		switch op {
		case OpJSLooseEqual:
			return OpJSJumpIfLooseNotEq
		case OpJSLooseNotEqual:
			return OpJSJumpIfLooseEq
		case OpJSStrictEq:
			return OpJSJumpIfStrictNotEq
		case OpJSStrictNeq:
			return OpJSJumpIfStrictEq
		case OpJSLess:
			return OpJSJumpIfNotLess
		case OpJSGreater:
			return OpJSJumpIfLessEqual
		}
	}
	return OpHalt
}

// optimizeDeadConditionalJumpPass removes unreachable true-branches for compile-time
// false conditions in `OpJumpIfFalse` patterns by NOP-filling bytes up to jump target.
func (c *Compiler) optimizeDeadConditionalJumpPass() bool {
	if c == nil || len(c.bytecode) == 0 {
		return false
	}

	targets := collectJumpTargets(c.bytecode)
	changed := false

	for ip := 0; ip < len(c.bytecode); {
		op := OpCode(c.bytecode[ip])
		size := opcodeOperandSize(op, c.bytecode, ip)
		instrEnd := ip + 1 + size
		if instrEnd > len(c.bytecode) {
			break
		}

		if op != OpJumpIfFalse {
			ip = instrEnd
			continue
		}

		target := int(binary.BigEndian.Uint32(c.bytecode[ip+1 : ip+5]))
		if target <= instrEnd || target > len(c.bytecode) || target <= ip {
			ip = instrEnd
			continue
		}

		condStart := findPreviousInstructionStart(c.bytecode, ip)
		for condStart >= 0 && OpCode(c.bytecode[condStart]) == OpNop {
			condStart = findPreviousInstructionStart(c.bytecode, condStart)
		}
		if condStart < 0 || OpCode(c.bytecode[condStart]) != OpConstant || condStart+3 > len(c.bytecode) {
			ip = instrEnd
			continue
		}

		constIdx := int(binary.BigEndian.Uint16(c.bytecode[condStart+1 : condStart+3]))
		if constIdx < 0 || constIdx >= len(c.constants) {
			ip = instrEnd
			continue
		}
		if !isCompileTimeFalseValue(c.constants[constIdx]) {
			ip = instrEnd
			continue
		}

		if hasTargetInRange(targets, instrEnd, target-1) {
			ip = instrEnd
			continue
		}

		mutated := false
		for p := instrEnd; p < target; p++ {
			if OpCode(c.bytecode[p]) != OpNop {
				c.bytecode[p] = byte(OpNop)
				mutated = true
			}
		}
		if mutated {
			changed = true
		}
		ip = target
	}

	return changed
}

func isCompileTimeFalseValue(v Value) bool {
	switch v.Type {
	case VTBool:
		return v.Num == 0
	case VTInteger:
		return v.Num == 0
	case VTDouble:
		return v.Flt == 0
	case VTEmpty, VTNull:
		return true
	case VTJSUndefined:
		return true
	case VTString:
		return v.Str == ""
	case VTObject:
		return v.Num == 0
	default:
		return false
	}
}

func findPreviousInstructionStart(bytecode []byte, before int) int {
	if before <= 0 || len(bytecode) == 0 {
		return -1
	}
	prev := -1
	for ip := 0; ip < len(bytecode) && ip < before; {
		size := opcodeOperandSize(OpCode(bytecode[ip]), bytecode, ip)
		next := ip + 1 + size
		if next > before {
			break
		}
		prev = ip
		ip = next
	}
	return prev
}

type intStackValue struct {
	isInt      bool
	isConstInt bool
	constInt   int64
	producerIP int
}

func pushIntStack(stack *[]intStackValue, value intStackValue) {
	*stack = append(*stack, value)
}

func popIntStack(stack *[]intStackValue) intStackValue {
	if len(*stack) == 0 {
		return intStackValue{}
	}
	last := len(*stack) - 1
	v := (*stack)[last]
	*stack = (*stack)[:last]
	return v
}

func clearIntInference(locals map[uint16]bool, globals map[uint16]bool, stack *[]intStackValue) {
	clear(locals)
	clear(globals)
	*stack = (*stack)[:0]
}

func rewriteDivisorConstantToShift(constants *[]Value, bytecode []byte, rhs intStackValue) (int64, bool) {
	if !rhs.isConstInt || rhs.constInt <= 0 {
		return 0, false
	}
	divisor := uint64(rhs.constInt)
	if divisor == 0 || divisor&(divisor-1) != 0 {
		return 0, false
	}
	shift := int64(bits.TrailingZeros64(divisor))
	if rhs.producerIP < 0 || rhs.producerIP+2 >= len(bytecode) {
		return 0, false
	}
	if OpCode(bytecode[rhs.producerIP]) != OpConstant {
		return 0, false
	}
	newConstIdx := len(*constants)
	*constants = append(*constants, NewInteger(shift))
	binary.BigEndian.PutUint16(bytecode[rhs.producerIP+1:rhs.producerIP+3], uint16(newConstIdx))
	return shift, true
}

// optimizeIntegerArithmeticPass performs one linear integer-inference pass over
// bytecode and rewrites arithmetic opcodes to integer fast paths when safe.
func (c *Compiler) optimizeIntegerArithmeticPass() bool {
	if c == nil || len(c.bytecode) == 0 {
		return false
	}

	targets := collectJumpTargets(c.bytecode)
	knownIntLocals := make(map[uint16]bool)
	knownIntGlobals := make(map[uint16]bool)
	stack := make([]intStackValue, 0, 32)
	changed := false

	for ip := 0; ip < len(c.bytecode); {
		if _, boundary := targets[ip]; boundary {
			clearIntInference(knownIntLocals, knownIntGlobals, &stack)
		}

		op := OpCode(c.bytecode[ip])
		size := opcodeOperandSize(op, c.bytecode, ip)
		instrEnd := ip + 1 + size
		if instrEnd > len(c.bytecode) {
			break
		}

		switch op {
		case OpConstant:
			idx := int(binary.BigEndian.Uint16(c.bytecode[ip+1 : ip+3]))
			entry := intStackValue{producerIP: ip}
			if idx >= 0 && idx < len(c.constants) {
				v := c.constants[idx]
				if v.Type == VTInteger {
					entry.isInt = true
					entry.isConstInt = true
					entry.constInt = v.Num
				}
			}
			pushIntStack(&stack, entry)

		case OpGetLocal:
			idx := binary.BigEndian.Uint16(c.bytecode[ip+1 : ip+3])
			pushIntStack(&stack, intStackValue{isInt: knownIntLocals[idx], producerIP: ip})

		case OpGetGlobal:
			idx := binary.BigEndian.Uint16(c.bytecode[ip+1 : ip+3])
			pushIntStack(&stack, intStackValue{isInt: knownIntGlobals[idx], producerIP: ip})

		case OpCoerceToValue:
			// No type information change for numeric inference.

		case OpSetLocal, OpLetLocal:
			idx := binary.BigEndian.Uint16(c.bytecode[ip+1 : ip+3])
			rhs := popIntStack(&stack)
			if rhs.isInt {
				knownIntLocals[idx] = true
			} else {
				delete(knownIntLocals, idx)
			}

		case OpSetGlobal, OpLetGlobal:
			idx := binary.BigEndian.Uint16(c.bytecode[ip+1 : ip+3])
			rhs := popIntStack(&stack)
			if rhs.isInt {
				knownIntGlobals[idx] = true
			} else {
				delete(knownIntGlobals, idx)
			}

		case OpIncLocalInt, OpDecLocalInt:
			idx := binary.BigEndian.Uint16(c.bytecode[ip+1 : ip+3])
			knownIntLocals[idx] = true

		case OpIncGlobalInt, OpDecGlobalInt:
			idx := binary.BigEndian.Uint16(c.bytecode[ip+1 : ip+3])
			knownIntGlobals[idx] = true

		case OpForNextFastInt:
			varIdx := binary.BigEndian.Uint16(c.bytecode[ip+1 : ip+3])
			endIdx := binary.BigEndian.Uint16(c.bytecode[ip+3 : ip+5])
			knownIntLocals[varIdx] = true
			knownIntLocals[endIdx] = true
			stack = stack[:0]

		case OpForNextFastGlobalInt:
			varIdx := binary.BigEndian.Uint16(c.bytecode[ip+1 : ip+3])
			endIdx := binary.BigEndian.Uint16(c.bytecode[ip+3 : ip+5])
			knownIntGlobals[varIdx] = true
			knownIntGlobals[endIdx] = true
			stack = stack[:0]

		case OpAdd, OpSub, OpMul:
			rhs := popIntStack(&stack)
			lhs := popIntStack(&stack)
			res := intStackValue{producerIP: ip}
			if lhs.isInt && rhs.isInt {
				res.isInt = true
				switch op {
				case OpAdd:
					c.bytecode[ip] = byte(OpIAdd)
				case OpSub:
					c.bytecode[ip] = byte(OpISub)
				case OpMul:
					c.bytecode[ip] = byte(OpIMul)
				}
				changed = true
			}
			pushIntStack(&stack, res)

		case OpIDiv:
			rhs := popIntStack(&stack)
			lhs := popIntStack(&stack)
			res := intStackValue{producerIP: ip}
			if lhs.isInt && rhs.isInt {
				res.isInt = true
				if _, ok := rewriteDivisorConstantToShift(&c.constants, c.bytecode, rhs); ok {
					c.bytecode[ip] = byte(OpIRightShift)
					changed = true
				}
			}
			pushIntStack(&stack, res)

		case OpIAdd, OpISub, OpIMul, OpIRightShift:
			_ = popIntStack(&stack)
			_ = popIntStack(&stack)
			pushIntStack(&stack, intStackValue{isInt: true, producerIP: ip})

		case OpDiv, OpMod, OpPow, OpConcat,
			OpEq, OpNeq, OpLt, OpGt, OpLte, OpGte, OpIsRef, OpIsNotRef,
			OpAnd, OpOr, OpXor, OpEqv, OpImp:
			_ = popIntStack(&stack)
			_ = popIntStack(&stack)
			pushIntStack(&stack, intStackValue{producerIP: ip})

		case OpNeg, OpNot:
			a := popIntStack(&stack)
			pushIntStack(&stack, intStackValue{isInt: a.isInt, producerIP: ip})

		case OpPop, OpWrite:
			_ = popIntStack(&stack)

		case OpWriteN:
			count := int(binary.BigEndian.Uint16(c.bytecode[ip+1 : ip+3]))
			for i := 0; i < count; i++ {
				_ = popIntStack(&stack)
			}

		case OpCall, OpCallMember, OpCallBuiltin, OpArraySet, OpMemberGet, OpMemberSet, OpMemberSetSet,
			OpJump, OpJumpIfFalse, OpJumpIfTrue, OpGotoLabel,
			OpJSJump, OpJSJumpIfFalse, OpJSJumpIfTrue, OpJSTryEnter,
			OpJSBreak, OpJSContinue, OpJSForInCleanup, OpJSJumpIfLessFast,
			OpJSCall, OpJSCallMember, OpJSTailCall, OpJSTailCallMember, OpJSNew:
			clearIntInference(knownIntLocals, knownIntGlobals, &stack)
		}

		ip = instrEnd
	}

	return changed
}

// optimizeLocalCopyPropagationPass performs conservative local copy propagation
// within one basic block by rewriting OpGetLocal operands in-place.
//
// It tracks copies created by direct sequences:
//
//	OpGetLocal src  [OpNop...]  OpLetLocal dst
//
// then rewrites subsequent OpGetLocal dst to OpGetLocal src while the mapping
// remains valid. The map is invalidated at jump targets, control-flow edges,
// calls, and any write to a local slot, which keeps propagation confined to one
// linear region with no cross-branch assumptions.
func (c *Compiler) optimizeLocalCopyPropagationPass() bool {
	if len(c.bytecode) == 0 {
		return false
	}
	targets := collectJumpTargets(c.bytecode)
	alias := make(map[uint16]uint16)
	changed := false

	for ip := 0; ip < len(c.bytecode); {
		if _, boundary := targets[ip]; boundary {
			clear(alias)
		}

		op := OpCode(c.bytecode[ip])
		size := opcodeOperandSize(op, c.bytecode, ip)
		instrEnd := ip + 1 + size
		if instrEnd > len(c.bytecode) {
			break
		}

		switch op {
		case OpGetLocal:
			local := binary.BigEndian.Uint16(c.bytecode[ip+1 : ip+3])
			if src, ok := alias[local]; ok {
				if src != local {
					binary.BigEndian.PutUint16(c.bytecode[ip+1:ip+3], src)
					local = src
					changed = true
				}
			}

			// Detect direct copy pattern: OpGetLocal src [OpNop*] OpLetLocal dst.
			next := instrEnd
			for next < len(c.bytecode) && OpCode(c.bytecode[next]) == OpNop {
				next++
			}
			if next+2 < len(c.bytecode) && OpCode(c.bytecode[next]) == OpLetLocal {
				dst := binary.BigEndian.Uint16(c.bytecode[next+1 : next+3])
				resolved := local
				for {
					up, ok := alias[resolved]
					if !ok || up == resolved {
						break
					}
					resolved = up
				}
				if dst != resolved {
					alias[dst] = resolved
				} else {
					delete(alias, dst)
				}
			}

			// Eliminate one redundant load pattern:
			//   OpGetLocal X [OpNop*] OpGetLocal X [OpNop*] OpPop
			// where no jump target lands inside the removed bytes.
			nextLoad := instrEnd
			for nextLoad < len(c.bytecode) && OpCode(c.bytecode[nextLoad]) == OpNop {
				nextLoad++
			}
			if nextLoad+2 < len(c.bytecode) && OpCode(c.bytecode[nextLoad]) == OpGetLocal {
				other := binary.BigEndian.Uint16(c.bytecode[nextLoad+1 : nextLoad+3])
				if other == local {
					nextAfterSecond := nextLoad + 3
					for nextAfterSecond < len(c.bytecode) && OpCode(c.bytecode[nextAfterSecond]) == OpNop {
						nextAfterSecond++
					}
					if nextAfterSecond < len(c.bytecode) && OpCode(c.bytecode[nextAfterSecond]) == OpPop {
						if !hasTargetInRange(targets, nextLoad, nextAfterSecond) {
							for p := nextLoad; p < nextLoad+3; p++ {
								c.bytecode[p] = byte(OpNop)
							}
							c.bytecode[nextAfterSecond] = byte(OpNop)
							changed = true
						}
					}
				}
			}

		case OpGetGlobal:
			global := binary.BigEndian.Uint16(c.bytecode[ip+1 : ip+3])
			nextLoad := instrEnd
			for nextLoad < len(c.bytecode) && OpCode(c.bytecode[nextLoad]) == OpNop {
				nextLoad++
			}
			if nextLoad+2 < len(c.bytecode) && OpCode(c.bytecode[nextLoad]) == OpGetGlobal {
				other := binary.BigEndian.Uint16(c.bytecode[nextLoad+1 : nextLoad+3])
				if other == global {
					nextAfterSecond := nextLoad + 3
					for nextAfterSecond < len(c.bytecode) && OpCode(c.bytecode[nextAfterSecond]) == OpNop {
						nextAfterSecond++
					}
					if nextAfterSecond < len(c.bytecode) && OpCode(c.bytecode[nextAfterSecond]) == OpPop {
						if !hasTargetInRange(targets, nextLoad, nextAfterSecond) {
							for p := nextLoad; p < nextLoad+3; p++ {
								c.bytecode[p] = byte(OpNop)
							}
							c.bytecode[nextAfterSecond] = byte(OpNop)
							changed = true
						}
					}
				}
			}

		case OpLetLocal, OpSetLocal, OpIncLocalInt, OpDecLocalInt:
			written := binary.BigEndian.Uint16(c.bytecode[ip+1 : ip+3])
			delete(alias, written)
			for k, v := range alias {
				if v == written {
					delete(alias, k)
				}
			}

		case OpForNextFastInt:
			// OpForNextFastInt writes to varLocalIdx (bytes ip+1..ip+2) and is a
			// conditional backward jump, so both the written slot and all aliases must
			// be invalidated at this basic-block boundary.
			written := binary.BigEndian.Uint16(c.bytecode[ip+1 : ip+3])
			delete(alias, written)
			for k, v := range alias {
				if v == written {
					delete(alias, k)
				}
			}
			clear(alias)

		case OpForNextFastGlobalInt:
			clear(alias)

		case OpJump, OpJumpIfFalse, OpJumpIfTrue, OpGotoLabel,
			OpJSJump, OpJSJumpIfFalse, OpJSJumpIfTrue, OpJSTryEnter,
			OpJSBreak, OpJSContinue, OpJSForInCleanup,
			OpJSJumpIfLessFast,
			OpCall, OpCallMember, OpCallBuiltin, OpJSCall, OpJSCallMember, OpJSTailCall, OpJSTailCallMember, OpJSNew:
			clear(alias)
		}

		ip = instrEnd
	}

	return changed
}

// optimizePeepholePass performs one forward scan of c.bytecode looking for
// constant-pair binary operations to fold. It advances a sliding window
// looking for:
//
//	OpConstant[hi][lo]  [OpNop…]  OpConstant[hi][lo]  [OpNop…]  <foldableBinOp>
//
// OpNop bytes between instructions are skipped so that chained folds (e.g.
// "a"&"b"&"c" → "ab"&"c") collapse in a single pass even after earlier folds
// have introduced padding nops.
// Returns true if any instruction was folded (signals another pass needed).
func (c *Compiler) optimizePeepholePass() bool {
	if len(c.bytecode) < 7 { // minimum: OpConstant(3) + OpConstant(3) + BinOp(1)
		return false
	}

	// Build the set of absolute byte-offsets that are jump-target landing points.
	targets := collectJumpTargets(c.bytecode)
	changed := false

	for i := 0; i < len(c.bytecode); {
		// First instruction must be OpConstant.
		if OpCode(c.bytecode[i]) != OpConstant {
			i++
			continue
		}

		// Skip any OpNop padding to find the second instruction.
		j := i + 3
		for j < len(c.bytecode) && OpCode(c.bytecode[j]) == OpNop {
			j++
		}
		if j+3 > len(c.bytecode) || OpCode(c.bytecode[j]) != OpConstant {
			i++
			continue
		}

		// Skip any OpNop padding to find the binary op.
		k := j + 3
		for k < len(c.bytecode) && OpCode(c.bytecode[k]) == OpNop {
			k++
		}
		if k >= len(c.bytecode) || !isFoldableVBSBinaryOp(OpCode(c.bytecode[k])) {
			i++
			continue
		}

		// Safety: no jump target may land on bytes i+1 through k inclusive.
		// A well-formed program never jumps to an operand byte, but j and k
		// are opcode bytes that could legitimately be named targets.
		if hasTargetInRange(targets, i+1, k) {
			i++
			continue
		}

		// Read the two constant indices (big-endian uint16).
		idxA := int(binary.BigEndian.Uint16(c.bytecode[i+1:]))
		idxB := int(binary.BigEndian.Uint16(c.bytecode[j+1:]))
		binOp := OpCode(c.bytecode[k])
		if idxA >= len(c.constants) || idxB >= len(c.constants) {
			i++
			continue
		}

		// Attempt compile-time evaluation.
		result, ok := foldVBSBinaryOp(c.constants[idxA], c.constants[idxB], binOp)
		if !ok {
			i++
			continue
		}

		// Fold success: update first OpConstant to reference the result and
		// fill every byte from i+3 through k (inclusive) with OpNop so that
		// absolute jump offsets into this region remain valid.
		newIdx := c.addConstant(result)
		binary.BigEndian.PutUint16(c.bytecode[i+1:], uint16(newIdx))
		for p := i + 3; p <= k; p++ {
			c.bytecode[p] = byte(OpNop)
		}
		changed = true
		// Stay at i: the newly written OpConstant may chain with another
		// constant+op pair immediately following the nop block.
	}
	return changed
}

// collectJumpTargets scans bytecode and returns a set of every absolute byte
// offset that is named as a landing point by a VBScript jump instruction.
func collectJumpTargets(bytecode []byte) map[int]struct{} {
	targets := make(map[int]struct{})
	for ip := 0; ip < len(bytecode); {
		op := OpCode(bytecode[ip])
		size := opcodeOperandSize(op, bytecode, ip)
		ip++
		switch op {
		case OpJump, OpJumpIfFalse, OpJumpIfTrue, OpGotoLabel,
			OpJSJump, OpJSJumpIfFalse, OpJSJumpIfTrue, OpJSTryEnter,
			OpJSJumpIfNullish, OpJSJumpIfNotNullish, OpJSJumpIfNotUndefined,
			OpJSCase, OpJSDefault, OpJSBreak, OpJSContinue,
			OpJumpIfNotEq, OpJumpIfEq, OpJumpIfNotLt, OpJumpIfLte, OpJumpIfNotIs,
			OpJSJumpIfLooseNotEq, OpJSJumpIfLooseEq, OpJSJumpIfStrictNotEq, OpJSJumpIfStrictEq, OpJSJumpIfNotLess, OpJSJumpIfLessEqual:
			// 4-byte absolute target immediately follows the opcode.
			if ip+4 <= len(bytecode) {
				targets[int(binary.BigEndian.Uint32(bytecode[ip:]))] = struct{}{}
			}
		case OpForNextFastInt:
			// Body target sits at bytes 6-9 of the operand field
			// (after varLocalIdx(2), endLocalIdx(2), stepSign(1)).
			if ip+9 <= len(bytecode) {
				targets[int(binary.BigEndian.Uint32(bytecode[ip+5:]))] = struct{}{}
			}
		case OpForNextFastGlobalInt:
			// Body target sits at bytes 6-9 of the operand field
			// (after varGlobalIdx(2), endGlobalIdx(2), stepSign(1)).
			if ip+9 <= len(bytecode) {
				targets[int(binary.BigEndian.Uint32(bytecode[ip+5:]))] = struct{}{}
			}
		case OpJSJumpIfLessFast:
			// Exit target sits at bytes 5-8 of the operand field
			// (after nameConstIdx(2), limitConstIdx(2)).
			if ip+8 <= len(bytecode) {
				targets[int(binary.BigEndian.Uint32(bytecode[ip+4:]))] = struct{}{}
			}
		}
		ip += size
	}
	return targets
}

// hasTargetInRange reports whether any collected jump target falls in [from, to].
func hasTargetInRange(targets map[int]struct{}, from, to int) bool {
	for pos := from; pos <= to; pos++ {
		if _, ok := targets[pos]; ok {
			return true
		}
	}
	return false
}

// isFoldableVBSBinaryOp reports whether a given opcode can be folded over two
// compile-time constant Values.
func isFoldableVBSBinaryOp(op OpCode) bool {
	switch op {
	case OpAdd, OpSub, OpMul, OpDiv, OpIDiv, OpMod, OpConcat:
		return true
	}
	return false
}

// foldVBSBinaryOp evaluates a binary operation over two constant Values at
// compile time. Returns (result, true) on success, or (Value{}, false) if the
// operand types are not supported or the operation would cause division by zero.
func foldVBSBinaryOp(a, b Value, op OpCode) (Value, bool) {
	switch op {
	case OpConcat:
		// & always converts both sides to string before concatenating.
		sa, ok1 := vbsConstantToString(a)
		sb, ok2 := vbsConstantToString(b)
		if ok1 && ok2 {
			return NewString(sa + sb), true
		}
	case OpAdd:
		return foldVBSNumericOp(a, b,
			func(x, y int64) int64 { return x + y },
			func(x, y float64) float64 { return x + y })
	case OpSub:
		return foldVBSNumericOp(a, b,
			func(x, y int64) int64 { return x - y },
			func(x, y float64) float64 { return x - y })
	case OpMul:
		return foldVBSNumericOp(a, b,
			func(x, y int64) int64 { return x * y },
			func(x, y float64) float64 { return x * y })
	case OpDiv:
		// VBScript / always produces a Double.
		fa, oka := vbsConstantToFloat(a)
		fb, okb := vbsConstantToFloat(b)
		if oka && okb && fb != 0 {
			return NewDouble(fa / fb), true
		}
	case OpIDiv:
		// VBScript \ is integer division (truncates toward zero).
		if a.Type == VTInteger && b.Type == VTInteger && b.Num != 0 {
			return NewInteger(a.Num / b.Num), true
		}
	case OpMod:
		// Only fold positive integers to avoid sign-convention edge cases.
		if a.Type == VTInteger && b.Type == VTInteger && b.Num > 0 && a.Num >= 0 {
			return NewInteger(a.Num % b.Num), true
		}
		fa, oka := vbsConstantToFloat(a)
		fb, okb := vbsConstantToFloat(b)
		if oka && okb && fb != 0 {
			return NewDouble(math.Mod(fa, fb)), true
		}
	}
	return Value{}, false
}

// foldVBSNumericOp applies an arithmetic operation to two constant Values when
// both are numeric (VTInteger or VTDouble), promoting to Double when necessary.
func foldVBSNumericOp(a, b Value, intOp func(int64, int64) int64, fltOp func(float64, float64) float64) (Value, bool) {
	switch {
	case a.Type == VTInteger && b.Type == VTInteger:
		return NewInteger(intOp(a.Num, b.Num)), true
	case a.Type == VTDouble && b.Type == VTDouble:
		return NewDouble(fltOp(a.Flt, b.Flt)), true
	case a.Type == VTInteger && b.Type == VTDouble:
		return NewDouble(fltOp(float64(a.Num), b.Flt)), true
	case a.Type == VTDouble && b.Type == VTInteger:
		return NewDouble(fltOp(a.Flt, float64(b.Num))), true
	}
	return Value{}, false
}

// vbsConstantToString converts a compile-time constant Value to a string
// representation suitable for the & concatenation operator.
// Only VTString, VTInteger, and VTDouble are supported.
func vbsConstantToString(v Value) (string, bool) {
	switch v.Type {
	case VTString:
		return v.Str, true
	case VTInteger:
		return strconv.FormatInt(v.Num, 10), true
	case VTDouble:
		// Use %g to match VBScript's default numeric-to-string format.
		return strconv.FormatFloat(v.Flt, 'g', -1, 64), true
	}
	return "", false
}

// vbsConstantToFloat converts a compile-time constant numeric Value to float64.
func vbsConstantToFloat(v Value) (float64, bool) {
	switch v.Type {
	case VTInteger:
		return float64(v.Num), true
	case VTDouble:
		return v.Flt, true
	}
	return 0, false
}

func (c *Compiler) optimizeFusedLoadBranchPass() bool {
	if len(c.bytecode) < 8 {
		return false
	}
	targets := collectJumpTargets(c.bytecode)
	changed := false

	for i := 0; i < len(c.bytecode); {
		op := OpCode(c.bytecode[i])
		if op != OpGetLocal && op != OpGetGlobal && op != OpJSGetName {
			size := opcodeOperandSize(op, c.bytecode, i)
			i += 1 + size
			continue
		}

		// The load size is 1 (opcode) + 2 (operand) = 3 bytes
		loadSize := 3
		j := i + loadSize

		// Skip OpNop padding
		for j < len(c.bytecode) && OpCode(c.bytecode[j]) == OpNop {
			j++
		}

		if j+5 > len(c.bytecode) {
			i += loadSize
			continue
		}

		jumpOp := OpCode(c.bytecode[j])
		if jumpOp != OpJumpIfFalse && jumpOp != OpJSJumpIfFalse {
			i += loadSize
			continue
		}

		if hasTargetInRange(targets, i+1, j) {
			i += loadSize
			continue
		}

		var fusedOp ExtOpCode
		if op == OpGetLocal && jumpOp == OpJumpIfFalse {
			fusedOp = ExtOpJumpLocalIfFalse
		} else if op == OpGetGlobal && jumpOp == OpJumpIfFalse {
			fusedOp = ExtOpJumpGlobalIfFalse
		} else if op == OpJSGetName && jumpOp == OpJSJumpIfFalse {
			fusedOp = ExtOpJSJumpNameIfFalse
		} else {
			i += loadSize
			continue
		}

		idx := binary.BigEndian.Uint16(c.bytecode[i+1 : i+3])
		target := binary.BigEndian.Uint32(c.bytecode[j+1 : j+5])

		// Fused branch is 1 (Prefix) + 1 (ExtOp) + 2 + 4 = 8 bytes.
		c.bytecode[i] = byte(OpExtPrefix)
		c.bytecode[i+1] = byte(fusedOp)
		binary.BigEndian.PutUint16(c.bytecode[i+2:i+4], idx)
		binary.BigEndian.PutUint32(c.bytecode[i+4:i+8], target)

		for p := i + 8; p <= j+4; p++ {
			c.bytecode[p] = byte(OpNop)
		}

		changed = true
		i = j + 5
	}
	return changed
}

func (c *Compiler) optimizeInPlaceMathPass() bool {
	if len(c.bytecode) < 11 {
		return false
	}
	targets := collectJumpTargets(c.bytecode)
	changed := false

	for i := 0; i < len(c.bytecode); {
		op := OpCode(c.bytecode[i])
		if op != OpGetLocal && op != OpGetGlobal {
			size := opcodeOperandSize(op, c.bytecode, i)
			i += 1 + size
			continue
		}

		idx := binary.BigEndian.Uint16(c.bytecode[i+1 : i+3])
		loadSize := 3
		j := i + loadSize

		for j < len(c.bytecode) && OpCode(c.bytecode[j]) == OpNop {
			j++
		}

		if j+3 > len(c.bytecode) || OpCode(c.bytecode[j]) != OpConstant {
			i += loadSize
			continue
		}

		constIdx := binary.BigEndian.Uint16(c.bytecode[j+1 : j+3])
		constSize := 3
		k := j + constSize

		for k < len(c.bytecode) && OpCode(c.bytecode[k]) == OpNop {
			k++
		}

		if k+1 > len(c.bytecode) {
			i += loadSize
			continue
		}

		binOp := OpCode(c.bytecode[k])
		if binOp != OpAdd && binOp != OpSub && binOp != OpConcat {
			i += loadSize
			continue
		}
		binOpSize := 1
		l := k + binOpSize

		for l < len(c.bytecode) && OpCode(c.bytecode[l]) == OpNop {
			l++
		}

		if l+3 > len(c.bytecode) {
			i += loadSize
			continue
		}

		storeOp := OpCode(c.bytecode[l])
		if storeOp != OpSetLocal && storeOp != OpSetGlobal && storeOp != OpLetLocal && storeOp != OpLetGlobal {
			i += loadSize
			continue
		}

		storeIdx := binary.BigEndian.Uint16(c.bytecode[l+1 : l+3])
		if storeIdx != idx || ((op == OpGetLocal && (storeOp == OpSetGlobal || storeOp == OpLetGlobal)) || (op == OpGetGlobal && (storeOp == OpSetLocal || storeOp == OpLetLocal))) {
			i += loadSize
			continue
		}

		if hasTargetInRange(targets, i+1, l) {
			i += loadSize
			continue
		}

		var fusedOp ExtOpCode
		if op == OpGetLocal && binOp == OpAdd {
			fusedOp = ExtOpAddLocalConst
		} else if op == OpGetGlobal && binOp == OpSub {
			fusedOp = ExtOpSubGlobalConst
		} else if op == OpGetLocal && binOp == OpConcat {
			fusedOp = ExtOpConcatLocalConst
		} else {
			i += loadSize
			continue
		}

		// 1 byte OpExtPrefix + 1 byte ExtOpCode + 2 bytes Idx + 2 bytes ConstIdx = 6 bytes
		c.bytecode[i] = byte(OpExtPrefix)
		c.bytecode[i+1] = byte(fusedOp)
		binary.BigEndian.PutUint16(c.bytecode[i+2:i+4], idx)
		binary.BigEndian.PutUint16(c.bytecode[i+4:i+6], constIdx)

		for p := i + 6; p <= l+2; p++ {
			c.bytecode[p] = byte(OpNop)
		}

		changed = true
		i = l + 3
	}
	return changed
}

func (c *Compiler) optimizeConstantPoolingPass() bool {
	if len(c.bytecode) < 6 { // at least 2 OpConstants
		return false
	}
	targets := collectJumpTargets(c.bytecode)
	changed := false

	for i := 0; i < len(c.bytecode); {
		if OpCode(c.bytecode[i]) != OpConstant {
			size := opcodeOperandSize(OpCode(c.bytecode[i]), c.bytecode, i)
			i += 1 + size
			continue
		}

		// Find contiguous sequence of OpConstant (ignoring OpNops)
		var constIndices []uint16
		var opPositions []int
		var lastValidPos int

		curr := i
		for curr < len(c.bytecode) {
			if OpCode(c.bytecode[curr]) == OpNop {
				curr++
				continue
			}
			if OpCode(c.bytecode[curr]) != OpConstant {
				break
			}
			if curr+3 > len(c.bytecode) {
				break
			}
			// If not the first const, check for jump targets in between
			if len(opPositions) > 0 && hasTargetInRange(targets, lastValidPos+1, curr) {
				break
			}
			constIndices = append(constIndices, binary.BigEndian.Uint16(c.bytecode[curr+1:curr+3]))
			opPositions = append(opPositions, curr)
			lastValidPos = curr
			curr += 3
			if len(constIndices) == 4 { // Max pool size is 4
				break
			}
		}

		if len(constIndices) < 2 {
			i += 3
			continue
		}

		// Replace the first OpConstant with the pooled instruction
		firstPos := opPositions[0]
		lastPos := opPositions[len(opPositions)-1]

		c.bytecode[firstPos] = byte(OpExtPrefix)
		var bytesWritten int

		if len(constIndices) == 4 {
			c.bytecode[firstPos+1] = byte(ExtOpConstant4)
			binary.BigEndian.PutUint16(c.bytecode[firstPos+2:firstPos+4], constIndices[0])
			binary.BigEndian.PutUint16(c.bytecode[firstPos+4:firstPos+6], constIndices[1])
			binary.BigEndian.PutUint16(c.bytecode[firstPos+6:firstPos+8], constIndices[2])
			binary.BigEndian.PutUint16(c.bytecode[firstPos+8:firstPos+10], constIndices[3])
			bytesWritten = 10
		} else if len(constIndices) == 3 {
			c.bytecode[firstPos+1] = byte(ExtOpConstant3)
			binary.BigEndian.PutUint16(c.bytecode[firstPos+2:firstPos+4], constIndices[0])
			binary.BigEndian.PutUint16(c.bytecode[firstPos+4:firstPos+6], constIndices[1])
			binary.BigEndian.PutUint16(c.bytecode[firstPos+6:firstPos+8], constIndices[2])
			bytesWritten = 8
		} else { // 2 constants
			c.bytecode[firstPos+1] = byte(ExtOpConstant2)
			binary.BigEndian.PutUint16(c.bytecode[firstPos+2:firstPos+4], constIndices[0])
			binary.BigEndian.PutUint16(c.bytecode[firstPos+4:firstPos+6], constIndices[1])
			bytesWritten = 6
		}

		for p := firstPos + bytesWritten; p <= lastPos+2; p++ {
			c.bytecode[p] = byte(OpNop)
		}

		changed = true
		i = lastPos + 3
	}
	return changed
}
