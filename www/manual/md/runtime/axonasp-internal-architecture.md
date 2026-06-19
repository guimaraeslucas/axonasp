# Understand AxonASP Internal Architecture (VM, Compiler, and JavaScript Engine)

## Overview

AxonASP is a Classic ASP execution platform implemented in Go. It runs VBScript and JavaScript (with JScript compatibility) through one shared virtual machine core.

This page explains the runtime architecture in practical terms for developers coming from IIS, COM, and ASP object-model workflows, even if they are newer to Go.

## Use This Mental Model First

Treat AxonASP like this:

- Lexer and compiler replace the old script host parser layer.
- Bytecode VM replaces script host execution internals.
- `Server.CreateObject` remains the extension boundary, but native implementations are Go structs instead of COM DLLs.
- ASP intrinsic objects (`Request`, `Response`, `Server`, `Session`, `Application`, `Err`) are pre-registered native VM objects.

If you already understand Classic ASP request flow, the architecture will feel familiar. The main difference is that object dispatch and language execution are explicit, typed Go code.

## High-Level Execution Flow

For each request (or CLI/TUI run), AxonASP executes this pipeline:

1. Read ASP source and split script/content segments.
2. Compile script segments to VM bytecode.
3. Execute bytecode in a stack-based VM.
4. Route object/property/method access through native dispatch maps.
5. Write output through `Response` object semantics.

All runtime modes share the same core in `axonvm/`:

- HTTP server
- FastCGI server
- CLI
- MCP server
- Test suite

## Language Pipeline Differences

## VBScript Path (Single-Pass, No AST)

VBScript compilation is direct token-to-bytecode emission:

1. Lexer tokenizes script.
2. Compiler emits opcodes immediately.
3. VM executes opcodes against `Value` runtime.

Why this matters:

- Lower compile-time memory usage.
- Lower latency for dynamic script and `Eval` paths.
- Compatibility logic is encoded in emission and runtime coercion rules.

## JavaScript Path (AST + Bytecode)

JavaScript support uses an AST pipeline in `jscript/`:

1. JavaScript source is parsed into AST nodes.
2. Compiler lowers AST to AxonASP VM bytecode.
3. VM executes JavaScript opcodes using the same VM core and `Value` type system.

JavaScript runtime state is VM-local and explicit in `vm.go`, including:

- Function objects and lexical environments.
- Object property descriptors and shape/slot caches.
- RegExp, Promise, Generator, Proxy, Intl, ArrayBuffer, and module state containers.
- Async task queues (`setTimeout`, microtasks, `nextTick`, immediate queue).

Result: VBScript and JavaScript share one execution engine while preserving language-specific semantics.

## VM Runtime Model

## Stack Machine and Typed Values

The VM is stack-based (`StackSize = 4096`) and bytecode-driven.

Runtime values are stored in `Value` (tagged union style):

- `Type` tag (`VTEmpty`, `VTString`, `VTInteger`, `VTDouble`, `VTNativeObject`, and others).
- Primitive payload fields (`Num`, `Flt`, `Str`).
- Optional structured payload (arrays, names, references).

This avoids reflection-based dispatch and keeps hot paths predictable.

## Call Frames, Scope, and Error Mode

Each procedure call stores a call frame containing:

- Return instruction pointer.
- Stack/frame restoration metadata.
- Bound object context for class members.
- ByRef write-back mapping.
- `On Error Resume Next` state restoration.

This is how AxonASP preserves Classic ASP/VBScript error and ByRef behavior while using a modern VM core.

## Intrinsic Object Wiring

Intrinsic ASP objects are prebound as native object IDs and loaded into global slots.

Examples in VM startup include:

- `Response`
- `Request`
- `Server`
- `Session`
- `Application`
- `ObjectContext`
- `Err`

Session state persists under `temp/session` and application state is process-memory-backed.

## Native Object Mapping Pattern (Codebase-Accurate)

This section documents the actual VM wiring pattern used in `axonvm/vm.go`, so registration steps and snippets match real runtime behavior.

## 1. Add a typed map in VM state

Each native library has a dedicated map keyed by dynamic ID:

```go
g3dbItems          map[int64]*G3DB
msxmlServerItems   map[int64]*MsXML2ServerXMLHTTP
msxmlDOMItems      map[int64]*MsXML2DOMDocument
pdfItems           map[int64]*G3PDF
```

Pattern: one map per concrete object type, no generic reflection registry.

## 2. Initialize map in VM constructor

The map is allocated in `NewVM` with `make(...)`.

```go
g3dbItems:        make(map[int64]*G3DB),
msxmlServerItems: make(map[int64]*MsXML2ServerXMLHTTP),
```

## 3. Register ProgID in `Server.CreateObject`

In native `Server` dispatch, `CreateObject` normalizes `progID` and routes by `progIDKey`.

Typical forms used today:

```go
if progIDKey == "g3db" {
   return vm.newG3DBObject()
}

if progIDKey == "msxml2.serverxmlhttp" || progIDKey == "msxml2.xmlhttp" || progIDKey == "microsoft.xmlhttp" {
   obj := NewMsXML2ServerXMLHTTP(vm)
   id := vm.nextDynamicNativeID
   vm.nextDynamicNativeID++
   vm.msxmlServerItems[id] = obj
   return Value{Type: VTNativeObject, Num: id}
}
```

Key behavior:

- ProgID matching is normalized and case-insensitive by lowercase key.
- Object instance is stored in a typed map.
- Returned handle is always `VTNativeObject` with dynamic numeric ID.

## 4. Route method calls in `dispatchNativeCall`

Method calls resolve by object ID map membership:

```go
if g3dbObject, exists := vm.g3dbItems[objID]; exists {
   return g3dbObject.DispatchMethod(member, args)
}
if g3dbRS, exists := vm.g3dbResultSetItems[objID]; exists {
   return g3dbRS.DispatchMethod(member, args)
}
```

The same pattern is repeated for each library/object family.

## 5. Route property get in native member resolution

Property reads use the same typed-map check pattern:

```go
if g3dbObject, exists := vm.g3dbItems[target.Num]; exists {
   return g3dbObject.DispatchPropertyGet(member)
}
```

## 6. Route property set in native set handler

Writable native properties must be routed explicitly to `DispatchPropertySet` in the native member set path.

Pattern:

```go
if obj, exists := vm.someLibraryItems[target.Num]; exists {
   return obj.DispatchPropertySet(member, args)
}
```

No implicit reflection fallback is used for these native maps.

## Why This Pattern Is Used

This design keeps the runtime deterministic:

- O(1) object lookup via map key.
- No reflection overhead in hot member-call paths.
- Explicit control of method/property behavior.
- Easier compatibility auditing for Classic ASP semantics.

## JavaScript Support Details That Matter in Practice

If you are extending JavaScript behavior, these runtime facts are important:

- JavaScript execution uses VM-owned registries (`jsFunctionItems`, `jsEnvItems`, `jsPropertyItems`, and others) instead of generic interface containers.
- Module loading state is request-local (`jsModuleInstances`, `jsModuleLoading`) to avoid cross-request leakage.
- Promise/microtask execution uses explicit VM queues, not background global schedulers.
- Buffer and typed array backing memory is tracked in VM maps (`jsArrayBuffers`, `jsSharedArrayBuffers`) for deterministic ownership.
- Intl-related object state is modeled in dedicated maps for each Intl type.

For Microsoft-ecosystem developers: think of this as replacing hidden script engine internals with explicit, inspectable runtime tables.

## Performance and Memory Behavior

AxonASP performance strategy:

- Avoid reflection-heavy execution paths.
- Avoid generic boxed dispatch in VM hot loops.
- Reuse stable bytecode where possible.
- Keep type coercions explicit and localized.

Memory safety strategy:

- Request-scoped object maps bound to VM instance lifetime.
- Explicit cleanup for native objects that hold external resources.
- Deterministic release patterns in libraries that wrap OS or external handles.

## Extension Checklist (Native Library)

Use this checklist when adding a new library:

1. Add `lib_<name>.go` with concrete struct and typed dispatch methods.
2. Add `lib_<name>_disabled.go` with correct build tags and disabled behavior.
3. Add map field to VM struct and allocate in constructor.
4. Add ProgID route in `Server.CreateObject` path.
5. Add method route in `dispatchNativeCall`.
6. Add property-get route in native property resolution path.
7. Add property-set route when writable members exist.
8. Return explicit `Value` types and raise explicit errors for operational failures.
9. Add Go tests for behavior and compatibility edges.

## Internal Files to Start With

- `axonvm/vm.go` for VM state, intrinsic objects, CreateObject routing, and native dispatch.
- `axonvm/opcode.go` for opcode taxonomy.
- `axonvm/value.go` for runtime value representation.
- `axonvm/compiler*.go` for VBScript token-to-bytecode emission.
- `jscript/` for JavaScript parser and AST logic.
- `axonvm/lib_*.go` for native library implementation patterns.

## Design Principles

- Compatibility-first behavior for Classic ASP workloads.
- Single-pass direct compilation for VBScript.
- AST-driven compilation for JavaScript support.
- One shared VM runtime with typed native object dispatch.
- Explicit VM wiring over reflection for performance and maintainability.