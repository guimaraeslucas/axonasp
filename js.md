# 🚀 AXONASP: JSCRIPT MODERNIZATION & ES6+ EXPANSION ROADMAP

This document serves as a high-precision checklist for implementing remaining ECMAScript 6 (ES6) and modern ES11-ES24 features into the AxonASP JScript engine.

## 🎯 CORE DIRECTIVES

1. **Strict Isolation:** Modify ONLY JScript-related files (`axonvm/compiler_jscript.go`, `axonvm/vm_jscript.go`, etc.). DO NOT touch VBScript logic or general VM state that could affect VBScript behavior. If you need to modify the VM, ensure it is strictly for JScript and does not introduce regressions or change the VBScript behavior.
2. **Performance Axioms:**
* **Zero-Allocation:** Avoid creating new Go objects on the heap during hot paths.
* **No Reflection:** Use the established `Value` struct and switch-based dispatch.
* **Minimal GC Impact:** Prefer native Go primitives and stack-based operations. Avoid the use of interfaces or any constructs that could trigger GC cycles.


3. **VM Architecture Context (Crucial):**
* The AxonASP Eval loop is procedural (a large `for` loop labeled `aspExecLoop`).
* It uses a custom memory-managed stack (`stack []Value`), a `callStack []CallFrame`, and `sp`, `fp`, and `ip` pointers.
* **NO Go Host Recursion:** User scripts run 100% isolated within the loop. Function calls (`OpCall`) just push a frame and jump `ip`. `OpRet` pops the frame and restores `ip`/`sp`/`fp`. Native Go recursion is strictly for native built-ins. Leverage this architecture heavily, especially for stack management and state pausing.


4. **Validation:** Every step MUST be accompanied by a GoLang test case in `axonvm/jscript_es6_test.go` and a javascript ASP test page in `./www/tests/test_*.asp` that must run with success in `axonasp-cli.exe -r <filename>`. Don't delete the test files, just add new ones for the new features. Ensure that all existing tests pass without modification to confirm no regressions.
5. After implementing the features, update the documentation in `./www/manual/md/javascript/jscript-es6-support.md` to reflect the new capabilities and any limitations.
6. Please think and do your best job. I trust you.

**Objective:** Implement missing ECMAScript 2020+ features, ergonomic APIs, and advanced architectural updates into the AxonASP JScript engine.

---

## 🛠️ PHASE 3: WELL-KNOWN SYMBOLS & PROTOCOLS (MEDIUM COMPLEXITY)

**Goal:** Deeply integrate Well-Known Symbols into the engine's built-in methods to ensure full ES6 behavioral compliance.

### Tasks:

* SUBPHASE 3.1: Array Protocol Hooks
* [ ] **`Symbol.isConcatSpreadable`:** Modify `Array.prototype.concat`. Before treating an argument as an array to be flattened, check for the `Symbol.isConcatSpreadable` property. Ensure this lookup is fast.
* [ ] **`Symbol.species`:** Update `Array.prototype.slice`, `map`, and `filter`. Instead of hardcoding `new Array()`, they must look up the constructor's `Symbol.species` to support subclassing.


* SUBPHASE 3.2: Runtime Hooks
* [ ] **`Symbol.hasInstance`:** Update the `OpJSInstanceOf` opcode logic. Before doing standard prototype chain traversal, check if the right-hand-side object has a `[Symbol.hasInstance]` method and invoke it.
* [ ] **`Symbol.unscopables`:** Update `OpWithEnter` / `OpJSGetName` logic. When resolving variables inside a `with` statement, check the object's `Symbol.unscopables` property to block specific property leakage.



---

## 🛠️ PHASE 4: UNICODE & FULL PLANE SUPPORT (MEDIUM TO HIGH COMPLEXITY)

**Goal:** Bring JScript into full compliance with the Unicode Standard, moving beyond the Basic Multilingual Plane (BMP).

### Tasks:

* SUBPHASE 4.1: Regular Expressions & Strings
* [ ] **`u` Flag in RegExp:** Update the regular expression engine compilation step. Support the `u` flag, enabling Unicode code point escapes (`\u{1F600}`) and Unicode property escapes.
* [ ] **String Code Points:** Implement `String.prototype.codePointAt()` and `String.fromCodePoint()`. Be extremely careful with Go's `rune` vs JS string length (UTF-16 surrogate pairs) to avoid out-of-bounds panics.


* SUBPHASE 4.2: Lexical Unicode Identifiers
* [ ] **AST & Lexer Updates:** Modify the Lexer in `./jscript/` to allow valid Unicode characters in variable/identifier names as defined by the ES6 spec. Ensure constants map correctly without encoding corruption in the single-pass compiler context.



---

## 🛠️ PHASE 5: PROXY/REFLECT INVARIANTS (HIGH COMPLEXITY)

**Goal:** Enforce the strict "Invariants of the Essential Internal Methods" for Proxies to prevent malicious or broken traps from corrupting the VM state.

### Tasks:

* SUBPHASE 5.1: Trap Validation Engine
* [ ] **Constraint Checking Logic:** Create a centralized validation mechanism inside `./jscript/` for Proxy traps.
* [ ] **Missing Property Traps:** Enforce that a Proxy cannot report a property as missing (`[[Get]]` returns undefined or `[[Has]]` returns false) if the target object has that property explicitly marked as non-configurable.


* SUBPHASE 5.2: Prototype & Extensibility Safety
* [ ] **Prototype Traps:** Ensure the `getPrototypeOf` trap cannot return a different prototype if the target object is non-extensible.
* [ ] **Error Handling:** If an invariant is broken, throw a precise `TypeError` using the `jscripterrorcodes.go` system. Test extensively with edge cases.



---

## 🛠️ PHASE 6: INTERNATIONALIZATION - THE `Intl` OBJECT (EXTREME COMPLEXITY)

**Goal:** Implement the complex `Intl` API for ES6 localization support, balancing spec compliance with our "Zero-Allocation" and memory constraint axioms.

### Tasks:

* SUBPHASE 6.1: Core Infrastructure
* [ ] **Locale Resolution:** Implement the default locale detection and fallback mechanisms required by the spec.
* [ ] **Global `Intl` Namespace:** Register the `Intl` object and stub out its constructors globally.


* SUBPHASE 6.2: Formatters
* [ ] **`Intl.DateTimeFormat`:** Wrap Go's `time` package logic to mirror JS locale strings. Map JS format options (e.g., `year: 'numeric', month: 'short'`) to Go format strings efficiently.
* [ ] **`Intl.NumberFormat`:** Implement localized number formatting. Be wary of memory allocations when dealing with dynamic string building.
* [ ] **`Intl.Collator`:** Implement localized string comparison. Rely on Go standard libraries where possible, but ensure strict adherence to ES-spec sorting rules.



---

## ✅ FINAL CHECKLIST FOR AGENT

1. **Gofmt:** Did you run `gofmt` on all modified files?
2. **VBScript Check:** Run `go test ./axonvm -run TestVBScript` to ensure zero regressions in the primary engine.
3. **Memory Profile:** Use `go test -bench` to ensure no new allocations were introduced in the JScript execution path (especially in standard library additions).
4. **Error Codes:** Did you use the correct error codes from `jscripterrorcodes.go` for syntax/runtime failures, especially regarding Proxy invariants?
5. **Branding:** Ensure all new files follow the G3pix copyright header format.
6. **Documentation:** Did you update `jscript-es6-support.md` with the newly supported APIs (e.g., `Intl`, `Math` expansions) and their limitations?
7. **Testing:** Did you add comprehensive test cases for each new feature in both Go and ASP test files? (Crucial for verifying proxy invariants and unicode parsing).
8. **Code Review:** Before finalizing, verify that `OpCall` limits and stack states are safely preserved when dealing with nested Symbol and Proxy runtime hooks.