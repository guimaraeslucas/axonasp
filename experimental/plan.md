# Incremental Plan: AST to Bytecode VM

This plan outlines the steps to migrate AxonASP from an AST-walking interpreter to a Bytecode Virtual Machine. Don't modify the AST interpreter or the main interpreter. This is a second way to interpret our asp code.

## Phase 1: Foundation (Current Step)
- [x] Analyze current AST.
- [x] Define minimal Bytecode Instruction Set (Opcodes).
- [x] Create VM data structures (Stack, Frames, CallStack).
- [x] Create a basic Compiler (AST -> Bytecode) for simple expressions (arithmetic, literals).
- [x] Create a basic VM loop to execute the simple bytecode.

## Phase 2: Core Language Features
- [x] Implement Control Flow (If, Select Case, Loops).
- [x] Implement Variable Scoping (Globals vs Locals).
- [x] Implement Assignments and identifier resolution.

## Phase 3: Functions and Procedures
- [x] Implement `Sub` and `Function` compilation.
- [x] Implement explicit Empty type that is fully compatible with VBScript
- [x] Implement rounds floats before Mod
- [x] Implement `Call` opcode and Return values.
- [x] Handle arguments (ByVal semantics implemented; ByRef partially handled via objects later).

## Phase 4: Integration with AxonASP
- [x] Implement full support for 32-bit and 64-bit platforms (Go handles this natively).
- [x] Implement .env variable and main.go support to use the VM engine or keep using the current AST walker (Implemented via AXONASP_VM env var in asp_executor).
- [x] Map existing `asp/*` and `server/*` libraries (G3JSON, ADODB, etc.) to the VM if user select VM in .env (Mapped via HostEnvironment interface).
- [x] Implement `OP_CALL_EXTERNAL` or similar to bridge VM to Go host functions if user select VM in .env (Implemented via OP_CALL and BuiltinFunction).
- [x] Update `server/executor.go` to optionally use the VM instead of the AST walker if user select VM in .env.
- [x] Implement add idiv, notOp, concat, and toString helper functions in vm.go,
- [x] Implement needed compiler features like MemberExpression alongside FunctionDeclaration and ClassDeclaration for proper method invocation and anything that you judge necessary to run ASP script.
- [x] Implement missing handling for the AST NothingLiteral and the VM opcode for OP_NOTHING
- [x] Optimize "hot paths" (e.g., specific opcodes for common operations like `i = i + 1`).
- [x] Optimize variable lookups (resolve names to indices at compile time where possible).
- [x] Implement `ast.ClassDeclaration` and `experimental.BuiltinFunction` with `gob` due to interface encoding needs
- [x] Implement class and function compilation alongside updating the bytecode cache registrations. 
- [x] extend the VM to when something is not implemented, try to run from the AST walker


## Phase 5: Optimization & Caching
- [ ] Implement Bytecode caching (serialize `Bytecode` struct to disk/memory).
- [ ] extend the server's VM host adapter to set variables using the ExecutionContext's SetVariable method.
## Phase 6: Full Migration
- [ ] Run full test suite (`www/tests/*.asp`) against the VM.
