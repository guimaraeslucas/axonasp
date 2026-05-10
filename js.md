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


## 🛠️ PHASE 4: DATA STRUCTURES & SYMBOLS (MEDIUM-HIGH COMPLEXITY)

**Goal:** Implement memory-safe collections, low-level buffers, and internal engine symbols.

### Tasks:

* [x] **Well-Known Symbols:** Expand the existing `Symbol` support to include global symbols: `Symbol.iterator`, `Symbol.toStringTag`, `Symbol.species`, `Symbol.hasInstance`, and `Symbol.toPrimitive`, ensuring they are correctly wired and recognized by the engine and can be used in user scripts.
* [x] **Binary Data (Typed Arrays & DataView):** Implement `ArrayBuffer`, `DataView`, and typed arrays (`Uint8Array`, `Int32Array`, `Float64Array`, etc.) for high-performance I/O. This will require careful memory management to ensure that the underlying byte buffers are allocated and freed correctly without leaks. Consider using Go's `unsafe` package for efficient memory handling, but ensure that all operations are bounds-checked to prevent memory corruption.
* [ ] **Weak Collections (`WeakMap` & `WeakSet`):** Implement collections that do not prevent GC of their keys.
    * *ATTENTION:* Implementing `WeakMap` and `WeakSet` in Go is non-trivial. You may need to use a combination of `runtime.SetFinalizer` or careful weak-reference management. Ensure thoroughly tested memory safety to prevent leaks in long-running ASP applications.
* [ ] **Final checklist**: Did you followed the final checklist at the end of this document after implementing these features?

---

## 🛠️ PHASE 5: ITERATION PROTOCOL & DESTRUCTURING (HIGH COMPLEXITY)

**Goal:** Standardize iteration mechanics and support `const [a, b] = arr;` and `const {x, y} = obj;` Please proceed step-by-step. Start by implementing Sub-Phase 5.1 and mapping out the compiler AST unrolling strategy for Sub-Phase 5.2, then wait for confirmation before advancing to the next phase. Think carefully about how to represent the temporary RHS object on the execution stack during multi-variable assignments. Mark each sub-phase as complete in this document before moving to the next one.
**Memory Warning:** Be extremely careful with stack depth. Deeply nested destructuring can exhaust the stack. Always test memory usage with deeply nested patterns (e.g., depth of 10) to ensure the VM handles it gracefully without a Go panic and without excessive heap allocations.

### Tasks:

* 🛠️ SUB-PHASE 5.1: The Iteration Protocol Foundation
    Before destructuring arrays, the VM must understand Iterables.
    * [x] **`Symbol.iterator`:** Ensure `Symbol.iterator` is implemented and globally available.
    * [x] **Built-in Iterators:** Implement native Go functions to return valid Iterator objects for `Array` and `String`. An Iterator object must have a `next()` method. Minimize heap allocations by reusing a single iterator struct for each type and resetting its state on each new iteration.
    * [x] **The Iterator Result:** The `next()` method must return a standard JScript object: `{ value: any, done: boolean }`. Ensure no unnecessary heap allocations occur when generating this struct in hot loops.
        * [x] **Testing:** Write tests to confirm that `for...of` loops and manual iterator usage work correctly with arrays and strings.


* 🛠️ SUB-PHASE 5.2: Object Destructuring (Property-Based) [x]
    Implement object destructuring first, as it does not rely on the iteration protocol.
    * [x] **Compiler (`ObjectPattern`):** Update `compileJScriptLexicalDeclaration` and `compileJScriptAssignment` to recognize `ObjectPattern`.
    * [x] **Null/Undefined Check:** The compiled bytecode MUST perform a check before attempting to read properties. If the RHS is `null` or `undefined`, it must throw a `TypeError` immediately, preventing any further property access attempts.
    * [x] **Nested Objects:** Handle nested patterns like `const { a: { b } } = obj` by chaining property loads recursively *during compilation*, emitting a flat sequence of opcodes that first loads `a`, checks it for null/undefined, then loads `b`.
    * [x] **Linear Unrolling:** For `const {x, y} = obj`, compile this into: Evaluate RHS -> Store in temporary VM register/stack -> Load `x` from temp -> Store in scope -> Load `y` from temp -> Store in scope. This approach avoids recursion in the VM and keeps memory usage predictable.
    * [x] **Null/Undefined Check:** The compiled bytecode MUST perform a check. If the RHS is `null` or `undefined`, it must throw a `TypeError` *before* attempting to read properties.
        * [x] **Testing:** Write tests for various object patterns, including nested destructuring and edge cases (e.g., missing properties, null/undefined RHS).

* 🛠️ SUB-PHASE 5.3: Array Destructuring (Iterator-Based)
    Now integrate the Iteration Protocol with destructuring.
    * [ ] **Compiler (`ArrayPattern`):** Update the compiler to recognize `ArrayPattern`.
    * [ ] **Execution Flow:** For `const [a, b] = iterable`, the compiler must emit instructions to:
        1. Get `Symbol.iterator` from the RHS.
        2. Call it to get the iterator object.
        3. Call `.next()` on the iterator. If `done: true`, assign `undefined` to `a`. Else, assign `value` to `a`.
        4. Repeat step 3 for `b`.
    * [ ] **Fallback:** Throw a `TypeError` if the RHS is not iterable (i.e., `[Symbol.iterator]` is undefined).
        * [ ] **Testing:** Write tests for array destructuring with various iterables (arrays, strings, custom iterables) and edge cases (e.g., non-iterable RHS).

* 🛠️ SUB-PHASE 5.4: Default Values & Rest Elements (`...`)
    * [ ] **Default Values:** Support `const {x = 10} = obj`. If the loaded property is strictly `undefined`, evaluate and assign the default value.
    * [ ] **Object Rest (`...rest`):** Create a new object containing all enumerable properties of the RHS *except* those explicitly destructured.
    * [ ] **Array Rest (`...rest`):** Call `.next()` continuously until `done: true`, pushing all yielded values into a new Array instance.
        * [ ] **Testing:** Write tests for default values and rest elements in both object and array destructuring contexts.
        
* 🛠️ SUB-PHASE 5.5: Validation & AST Memory Safety
    * [ ] **No Go Recursion in VM:** Verify that `vm_jscript.go` does not call itself to resolve nested patterns.
    * [ ] **Tests:** Write comprehensive tests in `axonvm/jscript_es6_test.go` and `./www/tests/test_destructuring.asp`. Test deeply nested destructuring (e.g., depth of 10) to ensure the VM stack handles it gracefully without a Go panic.
* [ ] **Final checklist**: Did you followed the final checklist at the end of this document after implementing these features?

---

## 🛠️ PHASE 6: ES6 CLASSES (HIGH COMPLEXITY)

**Goal:** Support `class C extends B { constructor() { super(); } method() {} }`

### Tasks:

* [ ] **Compiler Update:** Implement `compileJScriptClassDeclaration`.
* [ ] **Logic:** Classes are NOT hoisted, execute in Strict Mode, map `constructor` to a function, and map methods to the `.prototype`.
* [ ] **Super Binding:** Implement the `super` keyword by tracking the "Home Object" of methods to correctly resolve the prototype chain.
* [ ] **Final checklist**: Did you followed the final checklist at the end of this document after implementing these features?

---

## 🛠️ PHASE 7: PROXIES & REFLECTION (HIGH COMPLEXITY)

**Goal:** Introduce metaprogramming capabilities.

### Tasks:

* [ ] **Proxy Object:** Intercept fundamental operations (`get`, `set`, `apply`, `construct`). Requires deep hooks into the JScript member dispatch engine (`jsMemberGet`, `jsMemberSet`).
* [ ] **Reflect Object:** Expose the global `Reflect` API for programmatic object manipulation, ensuring parity with `Proxy` traps.
* [ ] **Final checklist**: Did you followed the final checklist at the end of this document after implementing these features?

---

## 🛠️ PHASE 8: STATE MACHINES (GENERATORS & ASYNC/AWAIT) (EXTREME COMPLEXITY)

**Goal:** Support pause/resume capabilities and asynchronous execution without blocking the ASP thread.

### Tasks:

* [ ] **Architectural Advantage:** Use the explicit `CallFrame`, `sp`, `fp`, and `ip` state array to your advantage. Pausing a generator means saving this exact state so it can be pushed back onto `vm.callStack` later.
* [ ] **State Machine Transformation:** The compiler must convert `function*` (`yield`) and `async` functions into resumable states.
* [ ] **Microtask Queue:** Implement a microtask queue in the VM that processes resolved promises before returning control to the ASP engine.
* [ ] **Constraint:** Ensure this does NOT interfere with the synchronous nature of VBScript or standard ASP objects (e.g., `Response.Write` must work correctly inside `yield` steps).
* [ ] **Final checklist**: Did you followed the final checklist at the end of this document after implementing these features?

---

## 🛠️ PHASE 9: ECMASCRIPT MODULES (EXTREME RISK)

**Goal:** Shift code loading architecture to support `import` / `export`.

### Tasks:

* [ ] **Dependency Resolution:** Create logic for loading and linking ES Modules. *Note: ASP is traditionally synchronous and based on `#include`.* This requires careful mapping to load ES Modules into isolated scope environments while maintaining the ASP lifecycle.
* [ ] **Module Caching:** Implement a caching mechanism to prevent reloading the same module multiple times.
* [ ] **Syntax & Semantics:** Update the parser to recognize `import` and `export` statements, and the compiler to handle module scope and bindings.
* [ ] **Testing:** This is a high-risk change. Ensure comprehensive and rigorous testing to prevent breaking existing ASP applications.
* [ ] **Final checklist**: Did you followed the final checklist at the end of this document after implementing these features?

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
9. **Check complete:** [x] Phase 5.2 Object Destructuring task is marked as complete in this file
