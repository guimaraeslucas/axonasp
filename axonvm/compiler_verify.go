package axonvm

import (
	"fmt"
	"os"
	"sort"
	"strings"
)

// Bytecode integrity checks for diagnosing a corrupted instruction stream.
//
// The symptom that motivated these: a VBScript page panicked inside the handler
// for OpJSSuperMemberGet, a JavaScript opcode, having read two unrelated bytes
// as a constant index. A VM executing an opcode the program never contained
// means the instruction pointer has drifted out of step with the instruction
// boundaries, and after that every byte it reads is meaningless.
//
// Drift is invisible at the point it happens and fatal thousands of
// instructions later, so these checks look for it directly, right after each
// optimizer pass, and name the pass that introduced it.
//
// Off unless AXONASP_VERIFY is set, because they walk the whole program.

var bytecodeVerifyEnabled = os.Getenv("AXONASP_VERIFY") != ""

// verifyBytecodeAlignment reports any jump whose target is not the first byte of
// an instruction.
//
// Every optimizer pass rewrites bytecode in place and pads with OpNop rather
// than deleting bytes, precisely so that absolute jump targets stay valid. A
// target that no longer lands on an instruction boundary means some pass wrote
// an instruction of a different length than the one it replaced - which is the
// one mistake that family of passes must never make.
func (c *Compiler) verifyBytecodeAlignment(tag string) []string {
	var problems []string

	boundaries := make(map[int]struct{}, len(c.bytecode)/3+1)
	ip := 0
	for ip < len(c.bytecode) {
		boundaries[ip] = struct{}{}
		op := OpCode(c.bytecode[ip])
		size := opcodeOperandSize(op, c.bytecode, ip)
		next := ip + 1 + size
		if next <= ip {
			problems = append(problems, fmt.Sprintf(
				"%s: %s at %d reports operand size %d, which would not advance",
				tag, op.String(), ip, size))
			break
		}
		if next > len(c.bytecode) {
			problems = append(problems, fmt.Sprintf(
				"%s: %s at %d claims %d operand bytes but only %d remain",
				tag, op.String(), ip, size, len(c.bytecode)-ip-1))
			break
		}
		ip = next
	}
	boundaries[len(c.bytecode)] = struct{}{} // a jump to the end is a clean halt

	targets := make([]int, 0, 16)
	for t := range collectJumpTargets(c.bytecode) {
		targets = append(targets, t)
	}
	sort.Ints(targets)

	for _, t := range targets {
		if t < 0 || t > len(c.bytecode) {
			problems = append(problems, fmt.Sprintf(
				"%s: jump target %d is outside the program (len %d)",
				tag, t, len(c.bytecode)))
			continue
		}
		if _, ok := boundaries[t]; ok {
			continue
		}
		problems = append(problems, fmt.Sprintf(
			"%s: jump target %d is mid-instruction%s",
			tag, t, describeNearest(c.bytecode, boundaries, t)))
	}
	return problems
}

// describeNearest names the instruction a stray target landed inside, which is
// usually enough to identify the pass that resized it.
func describeNearest(bytecode []byte, boundaries map[int]struct{}, target int) string {
	start := -1
	for p := target; p >= 0 && p > target-24; p-- {
		if _, ok := boundaries[p]; ok {
			start = p
			break
		}
	}
	if start < 0 {
		return ""
	}
	op := OpCode(bytecode[start])
	size := opcodeOperandSize(op, bytecode, start)
	return fmt.Sprintf(", %d bytes into %s at %d (which spans %d-%d)",
		target-start, op.String(), start, start, start+size)
}

// checkBytecode runs the integrity checks and reports anything they find.
//
// It reports rather than panics: one pass introducing drift is worth seeing
// alongside whatever the next pass then does with it, and a panic would hide
// everything after the first problem.
func (c *Compiler) checkBytecode(tag string) {
	if !bytecodeVerifyEnabled {
		return
	}
	problems := c.verifyBytecodeAlignment(tag)
	if len(problems) == 0 {
		return
	}
	limit := len(problems)
	if limit > 8 {
		limit = 8
	}
	msg := strings.Join(problems[:limit], "\n  ")
	if len(problems) > limit {
		msg += fmt.Sprintf("\n  ... and %d more", len(problems)-limit)
	}
	fmt.Fprintf(os.Stderr, "AXONASP BYTECODE PROBLEM\n  %s\n", msg)
}
