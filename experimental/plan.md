# Incremental Plan: AST to Bytecode VM

This plan outlines the steps to migrate AxonASP from an AST-walking interpreter to a Bytecode Virtual Machine.

## Phase 1: Foundation (Current Step)
- [x] Analyze current AST.
- [ ] Define minimal Bytecode Instruction Set (Opcodes).
- [ ] Create VM data structures (Stack, Frames, CallStack).
- [ ] Create a basic Compiler (AST -> Bytecode) for simple expressions (arithmetic, literals).
- [ ] Create a basic VM loop to execute the simple bytecode.

## Phase 2: Core Language Features
- [ ] Implement Control Flow (If, Select Case, Loops).
- [ ] Implement Variable Scoping (Globals vs Locals).
- [ ] Implement Assignments and identifier resolution.

## Phase 3: Functions and Procedures
- [ ] Implement `Sub` and `Function` compilation.
- [ ] Implement `Call` opcode and Return values.
- [ ] Handle arguments (ByVal vs ByRef semantics).

## Phase 4: Integration with AxonASP
- [ ] Map existing `asp/*` and `server/*` libraries (G3JSON, ADODB, etc.) to the VM.
- [ ] Implement `OP_CALL_EXTERNAL` or similar to bridge VM to Go host functions.
- [ ] Update `asp/asp_executor.go` to optionally use the VM instead of the AST walker.

## Phase 5: Optimization & Caching
- [ ] Implement Bytecode caching (serialize `Bytecode` struct to disk/memory).
- [ ] Optimize "hot paths" (e.g., specific opcodes for common operations like `i = i + 1`).
- [ ] Optimize variable lookups (resolve names to indices at compile time where possible).

## Phase 6: Full Migration
- [ ] Run full test suite (`www/tests/*.asp`) against the VM.
- [ ] Deprecate AST walker.
