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
 * made available under the same license terms.
 */
package axonvm

import (
	"testing"
)

// truncatedBytecodeCase is a hand-built program whose tail is cut mid-instruction, matching
// the bytecode shape that made the peephole walker index past the end of c.bytecode.
type truncatedBytecodeCase struct {
	name      string
	bytecode  []byte
	constants []Value
}

// truncatedBytecodeCases returns the truncation fixtures shared by the direct-invocation and
// compiler-walk regression tests. Every buffer is at least 7 bytes long, because
// optimizePeepholePass returns before any instruction decode below that size.
func truncatedBytecodeCases() []truncatedBytecodeCase {
	return []truncatedBytecodeCase{
		{
			// OpExtPrefix as the final byte: the extended opcode byte is missing.
			name:      "ext_prefix_cut_at_tail",
			bytecode:  []byte{byte(OpNop), byte(OpNop), byte(OpNop), byte(OpNop), byte(OpNop), byte(OpNop), byte(OpExtPrefix)},
			constants: []Value{NewInteger(1)},
		},
		{
			// OpJSObjectRest with only one of its two static-count bytes present.
			name:      "object_rest_cut_at_tail",
			bytecode:  []byte{byte(OpNop), byte(OpNop), byte(OpNop), byte(OpNop), byte(OpNop), byte(OpJSObjectRest), 0x01},
			constants: []Value{NewInteger(1)},
		},
		{
			// Constant pair followed by padding and a truncated escape instruction: the
			// fold window scans into the truncated tail before advancing.
			name: "constant_pair_before_truncated_ext_prefix",
			bytecode: []byte{
				byte(OpConstant), 0x00, 0x00,
				byte(OpConstant), 0x00, 0x00,
				byte(OpNop),
				byte(OpExtPrefix),
			},
			constants: []Value{NewInteger(1), NewInteger(2)},
		},
		{
			// Constant pair then a truncated OpJSObjectRest: the second window probe lands
			// exactly on the variable-length instruction.
			name: "constant_pair_before_truncated_object_rest",
			bytecode: []byte{
				byte(OpNop),
				byte(OpConstant), 0x00, 0x00,
				byte(OpConstant), 0x00, 0x00,
				byte(OpJSObjectRest), 0x01,
			},
			constants: []Value{NewInteger(1), NewInteger(2)},
		},
	}
}

// newTruncatedCompiler builds a Compiler backed by a private copy of the fixture bytecode so
// one pass mutating c.bytecode cannot influence the next assertion.
func newTruncatedCompiler(tc truncatedBytecodeCase) *Compiler {
	return &Compiler{
		bytecode:    append([]byte(nil), tc.bytecode...),
		constants:   append([]Value(nil), tc.constants...),
		constantMap: make(map[string]int, 8),
	}
}

// TestOpcodeOperandSizeTruncatedTailReturnsZero is the direct-invocation regression test for
// the index-out-of-range panic: a variable-length opcode whose operand bytes are absent must
// report 0 so a walker advances exactly one byte and terminates instead of slicing past the
// end of the buffer.
func TestOpcodeOperandSizeTruncatedTailReturnsZero(t *testing.T) {
	tests := []struct {
		name     string
		bytecode []byte
		ip       int
	}{
		// Reported case 1: bare OpExtPrefix, no extended opcode byte available.
		{name: "bare_ext_prefix", bytecode: []byte{byte(OpExtPrefix)}, ip: 0},
		{name: "ext_prefix_at_tail", bytecode: []byte{byte(OpNop), byte(OpNop), byte(OpExtPrefix)}, ip: 2},
		// Reported case 2: OpJSObjectRest with the static count low byte missing.
		{name: "object_rest_missing_count_low", bytecode: []byte{byte(OpJSObjectRest), 0x01}, ip: 0},
		{name: "object_rest_at_tail", bytecode: []byte{byte(OpNop), byte(OpNop), byte(OpJSObjectRest), 0x00}, ip: 2},
		// OpJSObjectRest with a count that promises more bytes than the buffer holds.
		{name: "object_rest_missing_static_keys", bytecode: []byte{byte(OpJSObjectRest), 0x00, 0x01}, ip: 0},
		{name: "object_rest_missing_dynamic_count", bytecode: []byte{byte(OpJSObjectRest), 0x00, 0x01, 0x00, 0x00}, ip: 0},
		{name: "object_rest_one_byte_short", bytecode: []byte{byte(OpJSObjectRest), 0x00, 0x01, 0x00, 0x00, 0x00}, ip: 0},
		{name: "object_rest_huge_count", bytecode: []byte{byte(OpJSObjectRest), 0xFF, 0xFF, 0x00, 0x00}, ip: 0},
		// Extended opcodes whose declared operand width does not fit in the buffer.
		{name: "ext_op_event_register_short", bytecode: []byte{byte(OpExtPrefix), byte(ExtOpRegisterClassEvent), 0x00}, ip: 0},
		{name: "ext_op_event_register_one_byte_short", bytecode: []byte{byte(OpExtPrefix), byte(ExtOpRegisterClassEvent), 0x00, 0x00, 0x00}, ip: 0},
		{name: "ext_op_jump_one_byte_short", bytecode: []byte{byte(OpExtPrefix), byte(ExtOpJumpGlobalIfFalse), 0x00, 0x00, 0x00, 0x00, 0x00}, ip: 0},
		{name: "ext_op_constant4_short", bytecode: []byte{byte(OpExtPrefix), byte(ExtOpConstant4), 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00}, ip: 0},
	}

	for _, tc := range tests {
		t.Run(tc.name, func(t *testing.T) {
			op := OpCode(tc.bytecode[tc.ip])
			got := opcodeOperandSize(op, tc.bytecode, tc.ip)
			if got != 0 {
				t.Fatalf("opcodeOperandSize(%s, len=%d, ip=%d) = %d, want 0 for a truncated buffer",
					op.String(), len(tc.bytecode), tc.ip, got)
			}
		})
	}
}

// TestOpcodeOperandSizeWellFormedWidths pins the operand widths of the variable-length
// families so a bounds guard can never silently shrink a valid instruction.
func TestOpcodeOperandSizeWellFormedWidths(t *testing.T) {
	tests := []struct {
		name     string
		bytecode []byte
		want     int
	}{
		{name: "object_rest_no_static_keys", bytecode: []byte{byte(OpJSObjectRest), 0x00, 0x00, 0x00, 0x00}, want: 4},
		{name: "object_rest_one_static_key", bytecode: []byte{byte(OpJSObjectRest), 0x00, 0x01, 0x00, 0x00, 0x00, 0x00}, want: 6},
		{name: "object_rest_two_static_keys", bytecode: []byte{byte(OpJSObjectRest), 0x00, 0x02, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00}, want: 8},
		{name: "ext_op_zero_operand", bytecode: []byte{byte(OpExtPrefix), byte(ExtOpFileOpen)}, want: 1},
		{name: "ext_op_zero_operand_js_write", bytecode: []byte{byte(OpExtPrefix), byte(ExtOpJSWrite)}, want: 1},
		{name: "ext_op_zero_operand_coerce_numeric", bytecode: []byte{byte(OpExtPrefix), byte(ExtOpCoerceToNumeric)}, want: 1},
		{name: "ext_op_arg_count", bytecode: []byte{byte(OpExtPrefix), byte(ExtOpFilePrint), 0x00, 0x00}, want: 3},
		{name: "ext_op_class_event", bytecode: []byte{byte(OpExtPrefix), byte(ExtOpRegisterClassEvent), 0x00, 0x00, 0x00, 0x00}, want: 5},
		{name: "ext_op_constant2", bytecode: []byte{byte(OpExtPrefix), byte(ExtOpConstant2), 0x00, 0x00, 0x00, 0x00}, want: 5},
		{name: "ext_op_constant3", bytecode: []byte{byte(OpExtPrefix), byte(ExtOpConstant3), 0x00, 0x00, 0x00, 0x00, 0x00, 0x00}, want: 7},
		{name: "ext_op_constant4", bytecode: []byte{byte(OpExtPrefix), byte(ExtOpConstant4), 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00}, want: 9},
		{name: "ext_op_jump_local_if_false", bytecode: []byte{byte(OpExtPrefix), byte(ExtOpJumpLocalIfFalse), 0x00, 0x00, 0x00, 0x00, 0x00, 0x00}, want: 7},
		{name: "primary_constant", bytecode: []byte{byte(OpConstant), 0x00, 0x01}, want: 2},
		{name: "primary_jump", bytecode: []byte{byte(OpJump), 0x00, 0x00, 0x00, 0x00}, want: 4},
		{name: "primary_call_member", bytecode: []byte{byte(OpCallMember), 0, 0, 0, 0, 0, 0, 0, 0}, want: 8},
		{name: "primary_register_class_field", bytecode: []byte{byte(OpRegisterClassField), 0, 0, 0, 0, 0, 0, 0, 0, 0}, want: 9},
		{name: "primary_nop", bytecode: []byte{byte(OpNop)}, want: 0},
		{name: "primary_super_index_get", bytecode: []byte{byte(OpJSSuperIndexGet)}, want: 0},
	}

	for _, tc := range tests {
		t.Run(tc.name, func(t *testing.T) {
			got := opcodeOperandSize(OpCode(tc.bytecode[0]), tc.bytecode, 0)
			if got != tc.want {
				t.Fatalf("opcodeOperandSize(%s, len=%d) = %d, want %d",
					OpCode(tc.bytecode[0]).String(), len(tc.bytecode), got, tc.want)
			}
		})
	}
}

// TestOpcodeOperandSizeUnknownExtOpDefaultWidth verifies the escape-opcode fallback width
// (used for extended instructions without a specialized operand layout) is still reported when
// its two operand bytes are present.
func TestOpcodeOperandSizeUnknownExtOpDefaultWidth(t *testing.T) {
	unknown, ok := firstUnknownExtOpcode()
	if !ok {
		t.Skip("every extended opcode byte has a specialized operand width")
	}
	bytecode := []byte{byte(OpExtPrefix), unknown, 0x00, 0x00}
	size := opcodeOperandSize(OpExtPrefix, bytecode, 0)
	if size != 3 {
		t.Skipf("extended opcode 0x%02X resolves to the specialized width %d; default path not exercised", unknown, size)
	}
	if got := opcodeOperandSize(OpExtPrefix, bytecode[:3], 0); got != 0 {
		t.Fatalf("opcodeOperandSize(OpExtPrefix, ext=0x%02X, truncated) = %d, want 0", unknown, got)
	}
}

// firstUnknownExtOpcode returns a byte value that ExtOpCode.String does not name, which means
// the extended dispatch in opcodeOperandSize routes it through the default width.
func firstUnknownExtOpcode() (byte, bool) {
	for i := range 256 {
		if ExtOpCode(i).String() == "ExtOpUnknown" {
			return byte(i), true
		}
	}
	return 0, false
}

// TestOpcodeOperandSizeNeverReadsPastBuffer sweeps every primary opcode byte against every
// truncation length and both a zero-filled and an 0xFF-filled tail. Any future opcode case
// that dereferences bytecode without a bounds guard makes this test panic, and any
// variable-length instruction that reports a width exceeding the buffer fails the invariant
// asserted below.
func TestOpcodeOperandSizeNeverReadsPastBuffer(t *testing.T) {
	fills := []byte{0x00, 0xFF}
	invariantOps := func(op OpCode) bool {
		return op == OpExtPrefix || op == OpJSObjectRest
	}

	for value := range 256 {
		for length := 1; length <= 16; length++ {
			for _, fill := range fills {
				bytecode := make([]byte, length)
				for i := range bytecode {
					bytecode[i] = fill
				}
				bytecode[0] = byte(value)

				for ip := 0; ip < length; ip++ {
					op := OpCode(bytecode[ip])
					size := opcodeOperandSize(op, bytecode, ip)
					if size < 0 {
						t.Fatalf("negative operand size %d for op 0x%02X at ip %d (len=%d, fill=0x%02X)",
							size, byte(bytecode[ip]), ip, length, fill)
					}
					if invariantOps(op) && size > 0 && ip+1+size > len(bytecode) {
						t.Fatalf("op %s at ip %d declares %d operand bytes past a %d-byte buffer (fill=0x%02X)",
							op.String(), ip, size, length, fill)
					}
				}
			}
		}
	}
}

// TestOptimizePeepholePassTruncatedBytecode is the compiler-walk regression test: the
// peephole pass decodes instructions through opcodeOperandSize, so a truncated tail used to
// crash the whole compile with an index-out-of-range panic.
func TestOptimizePeepholePassTruncatedBytecode(t *testing.T) {
	for _, tc := range truncatedBytecodeCases() {
		t.Run(tc.name, func(t *testing.T) {
			compiler := newTruncatedCompiler(tc)
			compiler.optimizePeepholePass()
		})
	}
}

// TestOptimizerPassesTruncatedBytecodeNoPanic runs every bytecode walker that shares the
// three-way opcode size contract over the truncation fixtures: each pass must terminate
// without panicking and without mutating a buffer length it cannot decode.
func TestOptimizerPassesTruncatedBytecodeNoPanic(t *testing.T) {
	passes := []struct {
		name string
		run  func(*Compiler) bool
	}{
		{name: "peephole", run: func(c *Compiler) bool { return c.optimizePeepholePass() }},
		{name: "local_copy_propagation", run: func(c *Compiler) bool { return c.optimizeLocalCopyPropagationPass() }},
		{name: "integer_arithmetic", run: func(c *Compiler) bool { return c.optimizeIntegerArithmeticPass() }},
		{name: "dead_conditional_jump", run: func(c *Compiler) bool { return c.optimizeDeadConditionalJumpPass() }},
		{name: "fused_branch", run: func(c *Compiler) bool { return c.optimizeFusedBranchPass() }},
		{name: "fused_load_branch", run: func(c *Compiler) bool { return c.optimizeFusedLoadBranchPass() }},
		{name: "in_place_math", run: func(c *Compiler) bool { return c.optimizeInPlaceMathPass() }},
		{name: "constant_pooling", run: func(c *Compiler) bool { return c.optimizeConstantPoolingPass() }},
	}

	for _, tc := range truncatedBytecodeCases() {
		for _, pass := range passes {
			t.Run(tc.name+"/"+pass.name, func(t *testing.T) {
				compiler := newTruncatedCompiler(tc)
				before := len(compiler.bytecode)
				pass.run(compiler)
				if len(compiler.bytecode) != before {
					t.Fatalf("pass %s changed bytecode length from %d to %d", pass.name, before, len(compiler.bytecode))
				}
				// collectJumpTargets shares the same decoder and runs first inside most
				// passes; assert it also survives the truncated tail.
				if targets := collectJumpTargets(compiler.bytecode); targets == nil {
					t.Fatalf("collectJumpTargets returned a nil set for %s", tc.name)
				}
			})
		}
	}
}
