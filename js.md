# 🚀 AXONASP: JSCRIPT MODERNIZATION & ES6+ EXPANSION ROADMAP

This document serves as a high-precision checklist for implementing ECMAScript 6 (ES6) and modern ES11-ES24 features into the AxonASP JScript engine.

## 🎯 CORE DIRECTIVES

1. **Strict Isolation:** Modify ONLY JScript-related files (`axonvm/compiler_jscript.go`, `axonvm/vm_jscript.go`, etc.). DO NOT touch VBScript logic or general VM state that could affect VBScript behavior. If you need to modify the VM, ensure it is strictly for JScript and does not introduce regressions or change the VBScript behavior.
2. **Performance Axioms:**
* **Zero-Allocation:** Avoid creating new Go objects on the heap during hot paths.
* **No Reflection:** Use the established `Value` struct and switch-based dispatch.
* **Minimal GC Impact:** Prefer native Go primitives and stack-based operations. Avoid the use of interfaces or any constructs that could trigger GC cycles.
3. **VM Architecture Context (Crucial):**
    * The AxonASP Eval loop is procedural (a large for loop labeled aspExecLoop).
    *It uses a custom memory-managed stack (stack []Value), a callStack []CallFrame, and sp, fp, and ip pointers.
* **NO Go Host Recursion:** User scripts run 100% isolated within the loop. Function calls (OpCall) just push a frame and jump ip. OpRet pops the frame and restores ip/sp/fp. Native Go recursion is strictly for native built-ins. Leverage this architecture heavily, especially for stack management and state pausing.
3. **Validation:** Every step MUST be accompanied by a GoLang test case in `axonvm/jscript_es6_test.go` and a javascript ASP test page in `./www/tests/test_*.asp` that must run with success in `axonasp-cli.exe -r <filename>`. Don't delete the test files, just add new ones for the new features. Ensure that all existing tests pass without modification to confirm no regressions.
4. After implementing the features, update the documentation in `./www/manual/md/javascript/jscript-es6-support.md` to reflect the new capabilities and any limitations.
5. Please think and do your best job. I trust you.

---

## 🛠️ PHASE 2: MODERN SYNTAX & OPERATORS (MEDIUM COMPLEXITY)

**Goal:** Implement ES11/2020 and ES12/2021 syntactic sugar. Requires updating the AST, parser, and emitting specific OpCodes or conditional jumps.

### Tasks:

* [ ] **Optional Chaining (`?.`):** Update the parser and compiler to safely short-circuit and return `undefined` instead of throwing an error when accessing properties on `null` or `undefined`.
* [ ] **Nullish Coalescing (`??`):** Implement the `??` operator to return the right-hand operand ONLY when the left-hand is strictly `null` or `undefined` (unlike `||` which checks falsy values).
* [ ] **Logical Assignment (`||=`, `&&=`, `??=`):** Implement these operators, ensuring the left side is only evaluated and reassigned if the logical condition is met.
* [ ] **Exponentiation Operator (`**`):** Implement the `**` operator for exponentiation, ensuring it correctly handles edge cases (e.g., negative bases, fractional exponents).
* [ ] **BigInt Support:** Implement the `n` suffix for BigInt literals and ensure arithmetic operations with BigInts are correctly handled without converting to regular numbers.
* [ ] **Test:** For each operator, write comprehensive test cases covering typical usage, edge cases, and error conditions (e.g., accessing properties on `null` with optional chaining should not throw an error).

---

## 🛠️ PHASE 3: TAIL CALL OPTIMIZATION (TCO) (MEDIUM COMPLEXITY)

**Goal:** Ensure function calls in the tail position do not increase the execution stack size. Don't allow the `stack []Value` to grow indefinitely with deep recursion. High risk if not implemented correctly, as it can lead to memory leaks or crashes.

* [ ] **Implementation Logic:** Thanks to the procedural `CallFrame` architecture, TCO is highly achievable. Instead of pushing a new `CallFrame`, identify if a call is a tail call. If so, replace the local variables/arguments on the current `stack []Value`, keep the current `CallFrame`, and simply jump the `ip` back to the start of the function's bytecode. This way, the stack size remains constant regardless of recursion depth. The compiler must emit a specific OpCode (e.g., OpTailCall) when it detects a tail call during compilation. The VM will then handle this OpCode by performing the TCO logic instead of a normal call. This is a critical optimization for functions that rely on recursion, such as those processing linked lists or performing mathematical computations.
* [ ] **Test:** Write a test with infinite or deep recursion in a tail-call position to verify it does not trigger a stack overflow or out-of-bounds slice error. This is a critical test to ensure the TCO implementation is correct and memory-safe.

---

## 🛠️ PHASE 4: DATA STRUCTURES & SYMBOLS (MEDIUM-HIGH COMPLEXITY)

**Goal:** Implement memory-safe collections, low-level buffers, and internal engine symbols.

### Tasks:

* [ ] **Well-Known Symbols:** Expand the existing `Symbol` support to include global symbols: `Symbol.iterator`, `Symbol.toStringTag`, `Symbol.species`, `Symbol.hasInstance`, and `Symbol.toPrimitive`.
* [ ] **Binary Data (Typed Arrays & DataView):** Implement `ArrayBuffer`, `DataView`, and typed arrays (`Uint8Array`, `Int32Array`, `Float64Array`, etc.) for high-performance I/O.
* [ ] **Weak Collections (`WeakMap` & `WeakSet`):** Implement collections that do not prevent GC of their keys.
    * *ATTENTION:* Implementing `WeakMap` and `WeakSet` in Go is non-trivial. You may need to use a combination of `runtime.SetFinalizer` or careful weak-reference management. Ensure thoroughly tested memory safety to prevent leaks in long-running ASP applications.

---

## 🛠️ PHASE 5: ITERATION PROTOCOL & DESTRUCTURING (HIGH COMPLEXITY)

**Goal:** Standardize iteration mechanics and support `const [a, b] = arr;` and `const {x, y} = obj;`

### Tasks:

* [ ] **Iteration Protocol (Iterators & Iterables):** Replace the current hardcoded `for...of` and spread operator logic. Support `Symbol.iterator` and implement `.next()` consumption so the engine can process any `{ value: any, done: boolean }` structure.
* [ ] **Compiler Update:** Refactor `compileJScriptLexicalDeclaration` and `compileJScriptAssignment` to handle `jsast.ArrayPattern` and `jsast.ObjectPattern`.
* [ ] **Recursive Unpacking:**
    * For **Arrays**: Use the Iteration Protocol, call `.next()`, and assign to identifiers. Handle rest elements `[a, ...rest]`.
    * For **Objects**: Iterate properties and assign members by name.

* [ ] **Memory Warning:** Be extremely careful with stack depth. Deeply nested destructuring can exhaust the stack.

---

## 🛠️ PHASE 6: ES6 CLASSES (HIGH COMPLEXITY)

**Goal:** Support `class C extends B { constructor() { super(); } method() {} }`

### Tasks:

* [ ] **Compiler Update:** Implement `compileJScriptClassDeclaration`.
* [ ] **Logic:** Classes are NOT hoisted, execute in Strict Mode, map `constructor` to a function, and map methods to the `.prototype`.
* [ ] **Super Binding:** Implement the `super` keyword by tracking the "Home Object" of methods to correctly resolve the prototype chain.

---

## 🛠️ PHASE 7: PROXIES & REFLECTION (HIGH COMPLEXITY)

**Goal:** Introduce metaprogramming capabilities.

### Tasks:

* [ ] **Proxy Object:** Intercept fundamental operations (`get`, `set`, `apply`, `construct`). Requires deep hooks into the JScript member dispatch engine (`jsMemberGet`, `jsMemberSet`).
* [ ] **Reflect Object:** Expose the global `Reflect` API for programmatic object manipulation, ensuring parity with `Proxy` traps.

---

## 🛠️ PHASE 8: STATE MACHINES (GENERATORS & ASYNC/AWAIT) (EXTREME COMPLEXITY)

**Goal:** Support pause/resume capabilities and asynchronous execution without blocking the ASP thread.

### Tasks:

* [ ] **Architectural Advantage:** Use the explicit `CallFrame`, `sp`, `fp`, and `ip` state array to your advantage. Pausing a generator means saving this exact state so it can be pushed back onto `vm.callStack` later.
* [ ] **State Machine Transformation:** The compiler must convert `function*` (`yield`) and `async` functions into resumable states.
* [ ] **Microtask Queue:** Implement a microtask queue in the VM that processes resolved promises before returning control to the ASP engine.
* [ ] **Constraint:** Ensure this does NOT interfere with the synchronous nature of VBScript or standard ASP objects (e.g., `Response.Write` must work correctly inside `yield` steps).

---

## 🛠️ PHASE 9: ECMASCRIPT MODULES (EXTREME RISK)

**Goal:** Shift code loading architecture to support `import` / `export`.

### Tasks:

* [ ] **Dependency Resolution:** Create logic for loading and linking ES Modules. *Note: ASP is traditionally synchronous and based on `#include`.* This requires careful mapping to load ES Modules into isolated scope environments while maintaining the ASP lifecycle.
* [ ] **Module Caching:** Implement a caching mechanism to prevent reloading the same module multiple times.
* [ ] **Syntax & Semantics:** Update the parser to recognize `import` and `export` statements, and the compiler to handle module scope and bindings.
* [ ] **Testing:** This is a high-risk change. Ensure comprehensive and rigorous testing to prevent breaking existing ASP applications.

---

## ✅ FINAL CHECKLIST FOR AGENT

1. **Gofmt:** Did you run `gofmt` on all modified files?
2. **VBScript Check:** Run `go test ./axonvm -run TestVBScript` to ensure zero regressions.
3. **Memory Profile:** Use `go test -bench` to ensure no new allocations were introduced in the JScript execution path.
4. **Error Codes:** Did you use the correct error codes from `jscripterrorcodes.go` for syntax/runtime failures?
5. **Branding:** Ensure all new files follow the G3pix copyright header format.
6. **Documentation:** Did you update `jscript-es6-support.md` with the new features and any limitations or known issues?
7. **Testing:** Did you add comprehensive test cases for each new feature in both Go and ASP test files?
8. **Code Review:** Before finalizing, review the code for any potential performance pitfalls, memory leaks, or edge cases that could arise from the new features.
9. **Check complete:** Ensure the required task is marked as complete in this file
