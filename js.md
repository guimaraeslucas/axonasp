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

**Objective:** Implement missing ECMAScript 2020+ features, ergonomic APIs, and advanced architectural updates into the AxonASP JScript engine.

---

Here is the comprehensive, phased prompt you can provide to your coding agent. It is strictly organized from the easiest, lowest-risk library additions to the most complex architectural and AST changes. It incorporates checkpoints for testing to ensure the VM does not break.

---

### Phase 4: AST and Compiler Overhauls (High Risk)

These features require modifications to `compiler_jscript.go` (Parser/Lexer) and `opcode.go`.

**Subphase 4.1: Private Class Fields (`#`)**

* **Target:** Class Fields and Private properties.
* **Implementation Tips:**
* **Lexer/Parser:** Update the AST to recognize the `#` prefix inside class bodies.
* **Compiler:** Emit a new hidden initialization sequence during `OpJSNewClass`.
* **VM (`vm_jscript.go`):** To maintain zero-allocation speed, do NOT use a separate dictionary for private fields. Instead, mangle the property name internally (e.g., map `#name` to a hidden string key like `\x00__priv_name`) so `OpJSMemberGet` can never accidentally access it via standard bracket notation `obj["#name"]`. Check lexical scope at compile-time to ensure validity.



**Subphase 4.2: Explicit Resource Management (`using`)**

* **Target:** `using` declarations and `Symbol.dispose`.
* **Implementation Tips:**
* **Parser:** Add support for `using varName = expr;`.
* **Compiler:** This is syntactic sugar for a `try...finally` block. Wrap the current lexical scope in an implicit `OpJSTryEnter`.
* **VM:** In the implicit `finally` block, emit opcodes to fetch `varName[Symbol.dispose]` and execute it (`OpJSCallMember`). Handle multiple `using` declarations by pushing them onto a stack and disposing in reverse order.


--

### Phase 5: Hard Constraints & Major Libraries (Extreme Complexity)

These are massive undertakings. Do not start these unless Phase 1-4 are 100% stable.

**Subphase 5.1: RegExp Engine Replacement (Optional/Deferred)**

* **Target:** Named Capture Groups, Lookbehind, Lookahead.
* **Implementation Tips:** Go's native `regexp` package guarantees linear time (O(n)) to prevent ReDoS attacks, which means it explicitly omits Lookaround and Backreferences. To support full JS RegExp, we would need to integrate a PCRE-compatible engine (like `regexp2`). Note: This breaks our strict "no external engines" rule, so advise the user before proceeding.

**Subphase 5.2: Temporal API**

* **Target:** The `Temporal` global object.
* **Implementation Tips:** This is a massive API. Implement it as a completely isolated native Go file (e.g., `lib_temporal.go`) that registers itself to the JScript global context. Avoid coupling its heavy date-math logic into the core VM loop.

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

## ✅ FINAL CHECKLIST FOR AGENT

1. **Gofmt:** Did you run `gofmt` on all modified files?
2. **VBScript Check:** Run `go test ./axonvm -run TestVBScript` to ensure zero regressions.
3. **Memory Profile:** Use `go test -bench` to ensure no new allocations were introduced in the JScript execution path.
4. **Error Codes:** Did you use the correct error codes from `jscripterrorcodes.go` for syntax/runtime failures?
5. **Branding:** Ensure all new files follow the G3pix copyright header format.
6. **Documentation:** Did you update `jscript-es6-support.md` with the new features and any limitations or known issues?
7. **Testing:** Did you add comprehensive test cases for each new feature in both Go and ASP test files?
8. **Code Review:** Before finalizing, review the code for any potential performance pitfalls, memory leaks, or edge cases that could arise from the new features.


