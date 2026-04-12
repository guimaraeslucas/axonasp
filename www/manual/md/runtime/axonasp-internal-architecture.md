# Understand AxonASP Internal Architecture (Go VM and Compiler)

## Overview

AxonASP is a modern execution engine for Classic ASP and VBScript, implemented in Go and designed for high compatibility, high throughput, and long-term maintainability.

This document explains how AxonASP works internally so that developers and automation agents can implement new runtime features, native libraries, and compatibility behaviors with confidence.

---

## Origin Story and Mission

AxonASP started from a practical risk: once VBScript deprecation was announced by Microsoft, long-term support for Classic ASP became uncertain.

The project mission is straightforward:
- Keep Classic ASP alive for production workloads.
- Preserve historical behavior where compatibility matters.
- Evolve the platform with modern capabilities without breaking legacy applications.

AxonASP modernizes the ecosystem by providing:
- Cross-platform execution on Windows, Linux, and macOS.
- Multiple runtime modes: HTTP server, FastCGI server, CLI interpreter, and MCP server.
- Modern developer workflows through CLI and TUI-style interactive execution, plus MCP integration for model-assisted development and web application maintenance.

---

## Core Engine Design

## High-Level Pipeline

AxonASP executes ASP/VBScript through a direct lexer-to-bytecode pipeline:

1. Source is tokenized by the lexer in `vbscript/`.
2. Compiler consumes tokens and emits VM opcodes directly.
3. Stack-based VM executes bytecode against a typed `Value` model.
4. Intrinsic ASP objects and native libraries are dispatched through VM-native object routing.

This architecture removes AST construction overhead and minimizes intermediate allocations.

## Runtime Modes Share One Core

The following executables reuse the same compiler/VM core in `axonvm/`:
- HTTP server (`server/`)
- FastCGI server (`fastcgi/`)
- CLI (`cli/`)
- MCP server (`mcp/`)

Feature parity is expected across all modes.

---

## Go Implementation Strategy and Optimizations

## Wrapping Go Libraries as ASP Objects

AxonASP extends Classic ASP by wrapping Go implementations behind `Server.CreateObject(...)` compatibility objects.

Typical flow:
- A ProgID (for example, `MSXML2.DOMDocument` or a custom library ProgID) is passed to `Server.CreateObject`.
- VM creates a concrete Go struct instance for that object.
- VM assigns a dynamic native ID and returns a `VTNativeObject` value.
- Member access is routed through typed dispatch functions, not reflection.

This allows ASP pages to consume modern Go functionality through object semantics that feel native to Classic ASP.

## Performance Discipline in Go

AxonASP prioritizes low-overhead execution:
- Avoid `interface{}`-centric runtime paths in VM hot loops.
- Avoid reflection (`reflect`) for object routing.
- Prefer explicit typed dispatch (`switch` / `strings.EqualFold`) for member resolution.
- Keep data in compact structs and primitive fields where possible.
- Reduce temporary allocations and unnecessary conversions in compile/execute phases.

The VM value model (`Value`) is a tagged struct that avoids generic boxed object payloads in the execution path.

---

## Compiler Architecture (No AST Rule)

## Rule: No AST

AxonASP does not build an Abstract Syntax Tree for script compilation.

Instead, the compiler performs immediate bytecode emission while reading tokens. This improves:
- Compilation latency
- Memory profile
- Throughput under dynamic script generation workloads

## What Replaces AST

The compiler uses:
- Token stream parsing with direct opcode emission.
- Constant pool indexing for names/literals.
- Jump patching for control-flow targets.
- Class/member registration opcodes for runtime metadata.
- Context-aware emission for VBScript compatibility semantics.

Representative behavior in compiler internals:
- Direct expression emission (`OpAdd`, `OpConcat`, comparisons, logical ops).
- Value-context coercion (`OpCoerceToValue`) where VBScript semantics require default-property reads.
- ByRef patching with argument reference opcodes (`OpArgGlobalRef`, `OpArgLocalRef`, `OpArgClassMemberRef`).
- Class/property registration opcodes (`OpRegisterClass*`) consumed by VM at runtime.

## Compiler Tooling for Runtime Efficiency

Compiler and execution tooling includes:
- Directive compilation support for `<%@ ... %>` blocks.
- Dynamic expression execution support (`Eval` / dynamic compile paths).
- Eval bytecode caching (LRU) keyed by expression and compare-mode/scope signatures.
- Script cache layers for reducing repeated compilation work.

The guiding pattern is stable bytecode reuse whenever scope and options permit.

---

## VM Architecture

## VM Type

AxonASP VM is:
- Stack-based
- Bytecode-driven
- Tagged-value runtime

Static stack capacity is fixed (`StackSize = 4096`) and execution is opcode-oriented.

## Opcode Model

The opcode set is organized by category:
- Data movement (`OpGetGlobal`, `OpSetLocal`, etc.)
- Arithmetic and logical operations
- Control flow and jumps
- Output writing (`OpWrite`, `OpWriteStatic`)
- Class registration and instantiation
- Member calls and dispatch (`OpCallMember`, `OpMemberGet`, `OpMemberSet`)
- Error mode handling (`OpOnErrorResumeNext`, `OpOnErrorGoto0`)

This keeps runtime deterministic and predictable for compatibility-critical behaviors.

## Call Frames and ByRef

User-defined Sub/Function calls create VM call frames that track:
- Return instruction pointer
- Frame pointers and stack restoration
- Bound object for class-member context
- ByRef write-back mapping to caller slots
- Error mode state restoration

This enables VBScript-compatible ByRef semantics without dynamic reflection.

---

## Intrinsic ASP Object Emulation

AxonASP pre-reserves intrinsic object slots and exposes them as native VM objects:
- `Response`
- `Request`
- `Server`
- `Session`
- `Application`
- `ObjectContext`
- `Err`

These are available through the ASP execution environment and participate in normal member dispatch paths.

State behavior:
- Session state persists under `temp/session` with ASP session cookies.
- Application state is process-resident memory.

Error information is bridged to the ASP error surface (`ASPError`) with line, column, file, and categorized runtime diagnostics.

---

## Type System and Memory Model

## Variant Representation in Go

VBScript dynamic values are represented by `axonvm.Value`:
- `Type` tag (`VTEmpty`, `VTNull`, `VTBool`, `VTInteger`, `VTDouble`, `VTString`, `VTDate`, `VTArray`, `VTObject`, `VTNativeObject`, `VTBuiltin`, `VTUserSub`, `VTArgRef`)
- Numeric fields (`Num`, `Flt`)
- String payload (`Str`)
- Array pointer (`Arr`)
- Name metadata (`Names`) for subroutines/object fields

This design supports:
- Fast type checks in hot opcodes
- Compatibility coercions
- Reference-aware call behavior (ByRef)
- Minimal object-shape ambiguity at runtime

## Compatibility Semantics

Core compatibility rules preserved by compiler + VM cooperation include:
- Case-insensitive identifiers/member names
- 1-based behaviors where required by VBScript semantics
- Property Get/Let/Set dispatch distinctions
- Object/value coercion behavior
- `On Error Resume Next` and recovery behavior

---

## Native Object and Library Lifecycle

## Creation

On `Server.CreateObject(progID)`:
1. VM normalizes ProgID for compatibility matching.
2. VM allocates concrete Go object instance.
3. VM assigns dynamic native ID.
4. Instance is stored in a VM-owned map for that library type.
5. Return value is `Value{Type: VTNativeObject, Num: dynamicID}`.

## Dispatch

Member calls are routed by VM through:
- Native ID range/type map resolution
- Library-specific `DispatchMethod` and `DispatchPropertyGet`/`DispatchPropertySet` logic
- Argument conversion and return value conversion using `Value` helpers

## Cleanup and Lifetime

Object lifetimes are tied to VM/request execution context and reference behavior. Deterministic cleanup patterns and map lifecycle management are critical to avoid retained objects after script completion.

---

## Practical Extension Guide

## Add a New Native Library (Recommended Pattern)

1. Create `axonvm/lib_<name>.go` with a concrete struct.
2. Implement strongly typed dispatch members:
   - `DispatchMethod(methodName string, args []Value) Value`
   - `DispatchPropertyGet(propertyName string) Value`
   - Optional property set handler when writable members exist.
3. Register object map in VM state.
4. Integrate ProgID creation in VM native call routing.
5. Route method/property dispatch for that dynamic ID space.
6. Use `Value` constructors (`NewString`, `NewInteger`, `NewBool`, etc.) consistently.
7. Raise standardized runtime errors for invalid args/state instead of silent `Empty` on operational failures.

## Add Compiler/VM Features Safely

When adding language features or compatibility behavior:
- Preserve no-AST direct emission architecture.
- Add opcodes only when necessary and wire full lifecycle (emit + execute + error handling).
- Keep compatibility semantics explicit in compiler emission rules.
- Verify parity in HTTP, FastCGI, CLI, and MCP execution paths.
- Add regression tests for compile and runtime behavior.

---

## Internal Files to Know First

Start here for architecture work:
- `axonvm/vm.go` - VM runtime, call frames, native object routing, intrinsic object plumbing.
- `axonvm/opcode.go` - Opcode definitions and instruction taxonomy.
- `axonvm/value.go` - Variant representation and value helpers.
- `axonvm/compiler*.go` - Token-to-bytecode direct emission logic.
- `vbscript/` - Lexer and VBScript token definitions/errors.
- `axonvm/lib_*.go` - Native library implementations.

---

## Design Principles Summary

AxonASP is built around four non-negotiable principles:
- Compatibility first for Classic ASP/VBScript behavior.
- Single-pass no-AST compilation for speed and memory efficiency.
- Stack-based bytecode VM with strict typed-value runtime.
- Native Go extensions exposed as ASP objects without reflection-heavy dispatch.

If you follow these principles, you can extend AxonASP quickly while preserving both performance and compatibility.