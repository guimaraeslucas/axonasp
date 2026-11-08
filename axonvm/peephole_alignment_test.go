package axonvm

import "testing"

// A peephole pass that advances a single byte from an OpConstant lands on that
// instruction's operand. The next iteration then reads the operand byte as an
// opcode, and when it happens to equal OpConstant the pass folds two unrelated
// byte pairs into a constant and writes the result over real instructions.
//
// The stream below is built so that a drifted scan finds a complete, foldable
// OpConstant/OpConstant/OpAdd sequence one byte off the real boundary, while a
// scan that advances properly never sees it. A correct pass leaves the bytecode
// untouched; the buggy one rewrites it.
//
// This needs no particular page to reproduce, but in real code it needed a very
// large one: the drifted byte has to collide with OpConstant and the fold has to
// succeed, so it surfaced on a 1.4MB program and never in a small test.
func TestPeepholePassNeverScansFromInsideAnOperand(t *testing.T) {
	c := &Compiler{
		constants: []Value{
			{Type: VTInteger, Num: 11},
			{Type: VTInteger, Num: 22},
		},
	}

	// Offset 0 is a real OpConstant. Its operand's high byte is itself
	// byte(OpConstant), which is what a one-byte advance would land on.
	//
	//   0: OpConstant  hi=OpConstant lo=0x00     <- real instruction
	//   3: 0x01                                  <- next real opcode, not OpConstant
	//   4: OpConstant  0x00 0x00                 <- only reachable by drifting
	//   7: OpAdd                                 <- completes the fake pattern
	//
	// Drifting to offset 1 reads a constant index from bytes 2-3, so those have
	// to name a real constant (1) or the pass rejects the index and bails out
	// before folding - which would make this test pass against the bug.
	c.bytecode = []byte{
		byte(OpConstant), byte(OpConstant), 0x00,
		0x01,
		byte(OpConstant), 0x00, 0x00,
		byte(OpAdd),
		byte(OpHalt),
	}
	before := append([]byte(nil), c.bytecode...)
	constantsBefore := len(c.constants)

	c.optimizePeepholePass()

	if got, want := len(c.constants), constantsBefore; got != want {
		t.Errorf("pass added %d constant(s); it folded something it never reached",
			got-want)
	}
	for i := range before {
		if c.bytecode[i] != before[i] {
			t.Fatalf("pass rewrote byte %d (%#02x -> %#02x): it scanned from "+
				"inside an operand and folded two unrelated byte pairs\n"+
				" before: %#v\n  after: %#v",
				i, before[i], c.bytecode[i], before, c.bytecode)
		}
	}
}

// The same invariant, stated directly: after optimization every jump target
// must still be the first byte of an instruction. Padding with OpNop instead of
// deleting bytes is what makes that true, so any pass that changes an
// instruction's length breaks it.
func TestOptimizedBytecodeKeepsJumpTargetsOnInstructionBoundaries(t *testing.T) {
	src := `
		dim i, total
		total = 0
		for i = 1 to 10
			if i mod 2 = 0 then
				total = total + i * 2 + 3
			else
				total = total + 1 + 2 + 4
			end if
		next
		response.write total
	`
	c := NewCompiler(src)
	if err := c.Compile(); err != nil {
		t.Fatalf("compile: %v", err)
	}
	if problems := c.verifyBytecodeAlignment("after compile"); len(problems) > 0 {
		for _, p := range problems {
			t.Errorf("%s", p)
		}
	}
}
