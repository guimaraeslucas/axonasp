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

## 🛠️ PHASE 7: PROXIES & REFLECTION (HIGH COMPLEXITY)

**Goal:** Introduce metaprogramming capabilities.

### Tasks:
Follow the subphase breakdown below for a structured implementation of Proxies and the Reflect API:
    * SUBPHASE 7.1: Core Types & Global Built-ins Setup
        * [ ] **Internal Representation:** Define the internal memory model for Proxies without breaking the `Value` struct. Either introduce a `VTJSProxy` type or utilize `VTJSObject` with hidden internal properties (e.g., `[[ProxyTarget]]` and `[[ProxyHandler]]`).
        * [ ] **Global Registration:** Inject the `Proxy` constructor and the `Reflect` namespace object into the global JScript environment upon VM initialization.
        * [ ] **Constructor Logic:** Implement the `new Proxy(target, handler)` built-in function. Ensure it throws a `TypeError` if `target` or `handler` are not valid objects (`VTJSObject` or `VTJSFunction`).
        * [ ] **Validation:** Create `test_proxy_init.asp` to verify `Proxy` and `Reflect` exist globally and that `new Proxy()` correctly validates its arguments.
    * SUBPHASE 7.2: Intercepting Property Access (`get` & `set` Traps)
        * [ ] **Get Trap:** Deeply hook into `vm.jsMemberGet`. If the object is a Proxy, inspect the `[[ProxyHandler]]` for a `"get"` property. If present, invoke it as a function with `(target, property, receiver)`. If not, forward the operation to the `[[ProxyTarget]]`.
        * [ ] **Set Trap:** Hook into `vm.jsMemberSet` and `vm.jsIndexSet`. Check the handler for a `"set"` property. Invoke it with `(target, property, value, receiver)`. 
        * [ ] **Strict Mode Enforcement:** In strict mode, if a `set` trap returns a falsy value, the VM MUST throw a `TypeError`.
        * [ ] **Validation:** Create `test_proxy_get_set.asp` to ensure properties can be dynamically intercepted, modified, or blocked without leaking memory or escaping the VM stack.
    * SUBPHASE 7.3: Intercepting Execution (`apply` & `construct` Traps)
        * [ ] **Callable Proxies:** A Proxy is only callable if its `[[ProxyTarget]]` is a `VTJSFunction`. Enforce this during instantiation.
        * [ ] **Apply Trap:** Hook into the VM's `OpCall` handler. If the callee is a Proxy, check for an `"apply"` trap. If present, invoke it with `(target, thisArg, argumentsList)`.
        * [ ] **Construct Trap:** Hook into the VM's `OpNew` handler. Check for a `"construct"` trap. Invoke it with `(target, argumentsList, newTarget)`. Ensure the return value is an object, otherwise throw a `TypeError`.
        * [ ] **Validation:** Create `test_proxy_apply_construct.asp` to test intercepting function calls and constructor invocations.
    * SUBPHASE 7.4: Intercepting Object Operations (`has`, `deleteProperty`, `ownKeys`)
        * [ ] **Has Trap:** Hook into the `in` operator logic (e.g., `OpJSIn`). Route to the `"has"` trap if defined.
        * [ ] **Delete Trap:** Hook into the `delete` operator logic. Route to the `"deleteProperty"` trap. Enforce strict mode throwing if the trap returns `false`.
        * [ ] **Keys/Enumeration:** Hook into `OpForIn` and `Object.keys()` internal logic to support the `"ownKeys"` trap, ensuring it returns a valid Array or iterable of strings/symbols.
        * [ ] **Validation:** Create `test_proxy_operations.asp` to verify operator interception works flawlessly.
    * SUBPHASE 7.5: The `Reflect` API Implementation
        * [ ] **Reflect Methods:** Implement `Reflect.get()`, `Reflect.set()`, `Reflect.apply()`, `Reflect.construct()`, `Reflect.has()`, `Reflect.deleteProperty()`, and `Reflect.ownKeys()`.
        * [ ] **Parity & Invocation:** Ensure these methods directly map to the internal VM dispatch mechanics (the exact same internal methods used when traps forward to the target).
        * [ ] **Return Semantics:** Unlike standard operators which might throw in strict mode, ensure `Reflect.set()` and `Reflect.deleteProperty()` return boolean success flags as dictated by the ES6 spec.
        * [ ] **Validation:** Create `test_reflect_api.asp` to verify parity between Proxy traps and Reflect invocations.
    * SUBPHASE 7.6: Final Agent Checklist
        * [ ] **Gofmt:** Run `gofmt` on all modified files.
        * [ ] **VBScript Check:** Run `go test ./axonvm -run TestVBScript` to ensure deep VM hooks into member resolution did NOT break VBScript `.` access.
        * [ ] **Memory Profile:** Run `go test -bench . -benchmem`. Proxy traps involve nested VM calls; ensure `CallFrame` allocations remain strictly stack-bound (Zero-Allocation axiom).
        * [ ] **Error Codes:** Ensure correct use of error codes from `jscripterrorcodes.go` for trap violations and TypeErrors.
        * [ ] **Documentation:** Update `jscript-es6-support.md` detailing the supported Proxy traps and the `Reflect` API features.

---

## 🛠️ PHASE 8: STATE MACHINES (GENERATORS & ASYNC/AWAIT) (EXTREME COMPLEXITY)
**Goal:** Support pause/resume capabilities and asynchronous execution without blocking the ASP thread and keeping the synchronous execution model intact. This is a critical phase that requires deep integration with the VM's execution model and careful handling of the microtask queue to ensure that asynchronous operations do not interfere with the synchronous nature of ASP. If coded wrong, it will create infinite loops or locks in the VM pool.

### Tasks:

*🛠️ SUBPHASE 8.1: MICROTASK QUEUE & PROMISES (MEDIUM COMPLEXITY)
**Goal:** Establish the foundational asynchronous primitives. Because our engine runs synchronously per HTTP request, Promises act as state containers, and the Microtask Queue is processed synchronously when the execution stack is empty or explicitly awaited.

* [x] **Native Promise Object:** Implement `Promise` in JScript (`resolve`, `reject`, `.then`, `.catch`, `.finally`).
* [x] **VM Microtask Queue:** Add a `jsMicrotaskQueue []func()` to the `VM` struct.
* [x] **Queue Processing:** Update the main `Run()` loop (or create a dedicated dispatcher) to pump/execute the `jsMicrotaskQueue` whenever the CallStack size reaches 0, or explicitly when an `await` instruction is hit.
* [x] **VM Reset Integrity:** Ensure `jsMicrotaskQueue` is perfectly cleared inside `resetDynamicMaps()` in `vm_pool.go` so queued tasks do not leak to the next user's HTTP request.

* [x] **Final checklist**: Did you followed the final checklist at the end of this document after implementing these features?

*🛠️ SUBPHASE 8.2: STATE MACHINES (GENERATORS & ASYNC/AWAIT) (HIGH COMPLEXITY)
**Goal:** Support pause/resume capabilities and asynchronous execution transparently.
* [x] **Check implementation**: ES6 Finally is a bit complex, it returns a promise that resolves with the original value after the finally callback finishes, check if your current implementation is following the ES6 finally specification.
* [x] **Compiler Transformation:** The compiler must convert `function*` (`yield`) and `async` functions into resumable state machines.
* [x] **Architectural Advantage (Pause/Resume):** Use the explicit `CallFrame`, `sp`, `fp`, and `ip` state array to your advantage. Pausing a generator or an `await` call simply means popping the current `CallFrame` and saving it into a closure or Promise structure, allowing the VM to execute other code or microtasks, and pushing it back onto `vm.callStack` to resume.
* [x] **The "Blocking Await" Shortcut:** Because we are in a dedicated goroutine, `await` does NOT need to yield to the Go scheduler. When `await` is called, the VM can simply loop and pump the Microtask queue until the specific Promise is resolved, blocking the execution synchronously but preserving strict ES6 semantic execution order.
* [x] **Constraint:** Ensure this does NOT interfere with the synchronous nature of VBScript or standard ASP objects (e.g., `Response.Write` must work correctly inside `yield` steps).

* [x] **Final checklist**: Did you followed the final checklist at the end of this document after implementing these features?

* 🛠️ SUBPHASE 8.3: ES MODULES (ESM) CACHE & REGISTRY (CRITICAL ARCHITECTURE)
**Goal:** Implement the split-caching architecture required to support `import` / `export` without destroying performance or leaking memory between concurrent ASP requests.

* [ ] **The Global AST Cache (Read-Only):** Modify the existing compiler/cache layer so that when an `import './module.js'` is encountered, the file is read and compiled into AST/Bytecode ONCE globally. Protect this shared cache using Go's `sync.RWMutex`.
* [ ] **The Request-Local Registry (Execution Memory):** Add a `jsModuleInstances map[string]*jsEnvFrame` (or similar environment pointer) to the `VM` struct. This represents the memory state of the modules *for the current user's request only*.
* [ ] **Singleton Emulation:** When a script calls `import`, the VM must check its request-local registry. If the module is not there, it retrieves the AST from the Global Cache, executes it to populate the variables, and stores the resulting environment in the local registry. If it is already there, it simply returns the existing environment reference.
* [ ] **VM Reset Integrity (CRITICAL):** Update `resetDynamicMaps()` in `vm_pool.go` to explicitly clear the request-local module registry (`clear(vm.jsModuleInstances)`). This guarantees that module state (e.g., `let loggedInUser = 'Admin';`) is destroyed when the request ends and does not leak to the next user's VM.


*🛠️ SUBPHASE 8.4: MODULE BINDING & EXECUTION (HIGH COMPLEXITY)

**Goal:** Connect the syntax to the new cache architecture.

* [ ] **AST & OpCodes:** Add support for `jsast.ImportDeclaration` and `jsast.ExportDeclaration`. Introduce specific OpCodes (e.g., `OpJSImport`, `OpJSExport`) to handle the resolution.
* [ ] **Synchronous Resolution:** Since the VM has a dedicated goroutine, module resolution (reading from disk if not in the global cache) must happen synchronously.
* [ ] **ASP Context Injection:** Ensure that standard ASP objects (`Response`, `Request`, `Session`) are automatically injected or accessible within the isolated Lexical Environment of the loaded module, preventing `ReferenceError` crashes when modules interact with the server.

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


