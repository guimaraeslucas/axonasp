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


## 🛠️ PHASE 6: PROXIES & REFLECTION (HIGH COMPLEXITY)

**Goal:** Introduce metaprogramming capabilities.

### Tasks:
Follow the subphase breakdown below for a structured implementation of Proxies and the Reflect API:
    * SUBPHASE 6.1: Core Types & Global Built-ins Setup
        * [x] **Internal Representation:** Define the internal memory model for Proxies without breaking the `Value` struct. Either introduce a `VTJSProxy` type or utilize `VTJSObject` with hidden internal properties (e.g., `[[ProxyTarget]]` and `[[ProxyHandler]]`).
        * [x] **Global Registration:** Inject the `Proxy` constructor and the `Reflect` namespace object into the global JScript environment upon VM initialization.
        * [x] **Constructor Logic:** Implement the `new Proxy(target, handler)` built-in function. Ensure it throws a `TypeError` if `target` or `handler` are not valid objects (`VTJSObject` or `VTJSFunction`).
        * [x] **Validation:** Create `test_proxy_init.asp` to verify `Proxy` and `Reflect` exist globally and that `new Proxy()` correctly validates its arguments.
    * SUBPHASE 6.2: Intercepting Property Access (`get` & `set` Traps)
        * [x] **Get Trap:** Deeply hook into `vm.jsMemberGet`. If the object is a Proxy, inspect the `[[ProxyHandler]]` for a `"get"` property. If present, invoke it as a function with `(target, property, receiver)`. If not, forward the operation to the `[[ProxyTarget]]`.
        * [x] **Set Trap:** Hook into `vm.jsMemberSet` and `vm.jsIndexSet`. Check the handler for a `"set"` property. Invoke it with `(target, property, value, receiver)`. 
        * [x] **Strict Mode Enforcement:** In strict mode, if a `set` trap returns a falsy value, the VM MUST throw a `TypeError`.
        * [x] **Validation:** Create `test_proxy_get_set.asp` to ensure properties can be dynamically intercepted, modified, or blocked without leaking memory or escaping the VM stack.
    * SUBPHASE 6.3: Intercepting Execution (`apply` & `construct` Traps)
        * [x] **Callable Proxies:** A Proxy is only callable if its `[[ProxyTarget]]` is a `VTJSFunction`. Enforce this during instantiation.
        * [x] **Apply Trap:** Hook into the VM's `OpCall` handler. If the callee is a Proxy, check for an `"apply"` trap. If present, invoke it with `(target, thisArg, argumentsList)`.
        * [x] **Construct Trap:** Hook into the VM's `OpNew` handler. Check for a `"construct"` trap. Invoke it with `(target, argumentsList, newTarget)`. Ensure the return value is an object, otherwise throw a `TypeError`.
        * [x] **Validation:** Create `test_proxy_apply_construct.asp` to test intercepting function calls and constructor invocations.
    * SUBPHASE 6.4: Intercepting Object Operations (`has`, `deleteProperty`, `ownKeys`)
        * [ ] **Has Trap:** Hook into the `in` operator logic (e.g., `OpJSIn`). Route to the `"has"` trap if defined.
        * [ ] **Delete Trap:** Hook into the `delete` operator logic. Route to the `"deleteProperty"` trap. Enforce strict mode throwing if the trap returns `false`.
        * [ ] **Keys/Enumeration:** Hook into `OpForIn` and `Object.keys()` internal logic to support the `"ownKeys"` trap, ensuring it returns a valid Array or iterable of strings/symbols.
        * [ ] **Proxy revocable:** Implement `Proxy.revocable` and any missing `Proxy` implementations.
        * [ ] **Object Traps:** Hook into `in` (`has`), `delete` (`deleteProperty`), and `Object.keys()` (`ownKeys`).
        * [ ] **Validation:** Create `test_proxy_operations.asp` to verify operator interception works flawlessly.
    * SUBPHASE 6.5: The `Reflect` API Implementation
        * [ ] **Reflect Methods:** Implement `Reflect.get()`, `Reflect.set()`, `Reflect.apply()`, `Reflect.construct()`, `Reflect.has()`, `Reflect.deleteProperty()`, and `Reflect.ownKeys()`.
        * [ ] **Parity & Invocation:** Ensure these methods directly map to the internal VM dispatch mechanics (the exact same internal methods used when traps forward to the target).
        * [ ] **Return Semantics:** Unlike standard operators which might throw in strict mode, ensure `Reflect.set()` and `Reflect.deleteProperty()` return boolean success flags as dictated by the ES6 spec.
        * [ ] **Validation:** Create `test_reflect_api.asp` to verify parity between Proxy traps and Reflect invocations.
    * SUBPHASE 6.6: Multi-Engine Execution Modes
        * [ ] **Configuration & State Management**: Read the engine mode configuration (`server.engine_mode`, `cli.engine_mode`, and `fastcgi.engine_mode`) from `axonasp.toml` using the existing `viper` implementation in `main.go` -> Inject this resolved mode value into the VM state so the parser knows how to handle the incoming request.
        * [ ] **Engine Mode Logic**: Implement conditional processing. If `default` -> process only `global.execute_as_asp` extensions as standard ASP files (parsing traditional `<% %>` delimiters). If `vbscript` -> process only `global.execute_as_vbscript` extensions as pure script, completely bypassing ASP delimiter parsing. If `javascript` -> process only `global.execute_as_javascript` extensions as pure script, completely bypassing ASP delimiter parsing.
        * [ ] **CLI Enhancements**: Implement a new CLI flag `-m` || `--mode` -> Allow the user to dynamically select `vbscript` or `javascript` modes from the command line (overriding the `.toml` configuration), and update the `--help` output to clearly display the new flag, its purpose, and accepted arguments. This option must work in the TUI mode and also in the run mode. 
        * [ ] **Documentation Updates**: Update ./www/manual/md/config/axonasp-toml.md -> Document the new configuration keys: server.engine_mode, cli.engine_mode, fastcgi.engine_mode, global.execute_as_vbscript, and global.execute_as_javascript. Update the runtime documentation in ./www/manual/md/runtime/cli-tui.md CLI/TUI section explaining the ability to run scripts interactively via a TUI or directly from the command line, and provide a comprehensive list of all CLI flags, ensuring the new -m || --mode flag is included.
    * SUBPHASE 6.7: Final Agent Checklist
        * [ ] **Gofmt:** Run `gofmt` on all modified files.
        * [ ] **JScript Check:** Run go tests on jscript implementation to ensure we're working as expected.
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


