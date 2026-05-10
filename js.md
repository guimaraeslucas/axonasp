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

## 🛠️ PHASE 6: ES6 CLASSES (HIGH COMPLEXITY)

**Goal:** Support `class C extends B { constructor() { super(); } method() {} }`


### Tasks:
Core Architecture Note for the Agent: Under the hood, ES6 Classes in JScript are "syntactic sugar" over JavaScript's existing prototype-based inheritance. However, they introduce strict semantics (e.g., they must be called with new, they are not hoisted, and all code inside them runs in Strict Mode). The implementation should leverage existing VTJSFunction and VTJSObject types, manipulating their internal properties (like __proto__ and prototype) via bytecode generation, rather than introducing entirely new Go structs. The compiler will need to generate bytecode that sets up the prototype chain correctly, handles the `super` keyword by tracking the "Home Object" of methods, and ensures that the constructor function is properly defined and linked to the class prototype. This will require careful management of the call stack and execution context to ensure that method calls and property accesses resolve correctly according to ES6 semantics.
    * SUBPHASE 6.1: Parser & AST Verification
        * [ ] **AST Nodes:** Verify or add AST nodes in `jscript/ast` for `ClassDeclaration`, `ClassExpression`, `MethodDefinition` (kinds: `constructor`, `method`, `get`, `set`), and `Super`.
        * [ ] **Parser Update:** Ensure the parser correctly handles the `class`, `extends`, and `super` keywords.
        * [ ] **Validation:** Add pure parser tests ensuring `class A extends B { constructor() { super(); } method() {} }` parses into the correct AST tree without VM execution.
    * SUBPHASE 6.2: Basic Class Compilation (The Constructor)
        * [ ] **Compiler Update:** Implement `compileJScriptClassDeclaration` in `axonvm/compiler_jscript.go`.
        * [ ] **TDZ Binding:** Treat the class declaration similarly to a `let` binding. Classes are NOT hoisted and must reside in the Temporal Dead Zone until evaluated.
        * [ ] **Constructor Logic:** Extract the `constructor` method and compile it as a standard `VTJSFunction`, tagging it internally with a flag (e.g., `IsClassConstructor: true`).
        * [ ] **VM Enforcement:** If a `VTJSFunction` tagged as a class constructor is invoked without the `new` operator (via `OpCall` instead of `OpNew`), the VM MUST throw a `TypeError` ("Class constructor cannot be invoked without 'new'").
        * [ ] **Validation:** Create `test_class_basic.asp` testing instantiation with `new` (success) and without `new` (throws TypeError).
    * SUBPHASE 6.3: Instance Methods & Strict Mode Enforcement
        * [ ] **Strict Mode:** Ensure the compiler sets the `StrictMode` flag for the constructor and all methods generated within the class block, as class bodies implicitly run in Strict Mode.
        * [ ] **Method Compilation:** Modify `compileJScriptClassDeclaration` to iterate over all `MethodDefinition` AST nodes where `static` is false, compiling each as a `VTJSFunction`.
        * [ ] **Prototype Binding:** Generate bytecode to assign these compiled methods to the constructor's `prototype` object (Equivalent to: `OpPush <Constructor>` -> `OpPush "prototype"` -> `OpMemberGet` -> `OpPush <MethodClosure>` -> `OpPush "<MethodName>"` -> `OpMemberSet`).
        * [ ] **Validation:** Create `test_class_methods.asp` to verify method calls, correct `this` context, and strict mode enforcement.
    * SUBPHASE 6.4: Static Methods and Accessors (Getters/Setters)
        * [ ] **Static Methods:** If a `MethodDefinition` is marked `static`, bind it directly to the constructor function object instead of its `prototype`.
        * [ ] **Accessors:** For `get` and `set` methods, use or implement the VM equivalent of `Object.defineProperty` using a specialized opcode (e.g., `OpJSDefineProperty`) to define descriptors on the `VTJSObject`.
        * [ ] **Validation:** Create `test_class_static_accessors.asp`. Test `Class.staticMethod()` and `instance.getterProp`.
    * SUBPHASE 6.5: Inheritance (`extends`) & Prototype Chaining
        * [ ] **SuperClass Evaluation:** If `extends <SuperClass>` is present, evaluate it. The evaluated `<SuperClass>` must be a valid constructor (`VTJSFunction`) or `null`. If not, throw a `TypeError`.
        * [ ] **Prototype Wiring:** Set the internal `__proto__` of the subclass's `prototype` object to `<SuperClass>.prototype`.
        * [ ] **Static Inheritance:** Set the internal `__proto__` of the subclass constructor itself to `<SuperClass>` to allow inheritance of static methods. Use internal VM assignments (e.g., `vm.jsSetProto`) to avoid Go heap allocation.
        * [ ] **Validation:** Create `test_class_extends.asp`. Check if instances inherit methods from the superclass prototype and if the subclass inherits static methods.
    * SUBPHASE 6.6: The `super()` Call in Constructors
        * [ ] **TDZ for `this`:** In a derived class constructor, `this` is uninitialized until `super()` is called. If `this` is accessed before `super()` completes, the VM must throw a `ReferenceError`.
        * [ ] **Super Invocation:** Compile `super(...)` as a special call that invokes the parent constructor using `Reflect.construct` logic via `OpNew`, explicitly targeting the extended class.
        * [ ] **Validation:** Create `test_class_super_constructor.asp`. Test constructor argument passing and verify `ReferenceError` if `this` is used prematurely.
    * SUBPHASE 6.7: The `super.method()` Call & Home Object Binding
        * [ ] **Home Object Assignment:** When attaching methods to the prototype, attach a hidden internal VM property `[[HomeObject]]` to the `VTJSFunction` pointing to the object it was bound to. Store this as an integer ID pointing to the VM's `jsObjectItems` array to strictly avoid reflection and Go heap escape.
        * [ ] **Super Resolution:** Translate `super.foo` to: Get `[[HomeObject]]` -> Get its prototype -> Lookup property `"foo"` -> Call with current `this` context.
        * [ ] **Validation:** Create `test_class_super_method.asp`. Test complex prototype chains (e.g., `C extends B extends A`).
    * SUBPHASE 6.8: Final Agent Checklist
        * [ ] **Gofmt:** Run `gofmt` on all modified files.
        * [ ] **VBScript Check:** Run `go test ./axonvm -run TestVBScript` to ensure strictly zero regressions on the VBScript side.
        * [ ] **Memory Profile:** Run `go test -bench . -benchmem` to guarantee no new allocations were introduced in the JScript execution path (Zero-Allocation axiom).
        * [ ] **Error Codes:** Ensure correct use of error codes from `jscripterrorcodes.go` for syntax/runtime failures.
        * [ ] **Documentation:** Update `jscript-es6-support.md` reflecting the new capabilities and any limitations.

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
9. **Check complete:** [x] Phase 5 Iteration Protocol & Destructuring task is marked as complete in this file
