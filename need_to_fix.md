# Architectural Plan: Compatibility & Engine Patches

## 7. Implicit ByRef Variable Allocation
- **Affected Files:**
  - `axonvm/compiler_expr.go` (`rewriteLastVarLoadToArgRef`)
  - `axonvm/vm.go` (`collectByRefsAndUnwrap`, `applyByRefWritebacksFromArgs`, `OpRet`)
- **Defect:**
  Implicitly declared variables passed ByRef to procedures returning objects/recordsets lose their reference because the compiler disables ByRef rewrite for undeclared variables (`!c.declaredGlobals[name]`), passing them ByVal. Additionally, `VTNativeObject` assignments are not written back properly to caller scopes.
- **Implementation Strategy:**
  1. In `compiler_expr.go` (`rewriteLastVarLoadToArgRef`):
     - Remove the restrictive check `if !c.declaredGlobals[name] { return }`.
     - If the variable is undeclared, allocate it dynamically in caller scope (registering as implicit global or local) and emit `OpArgGlobalRef` or `OpArgLocalRef`.
  2. In `axonvm/vm.go` (`OpRet` & `applyByRefWritebacksFromArgs`):
     - When callee updates a ByRef argument with a `VTNativeObject` (or `VTObject`), write the object value directly into the caller's target slot (`vm.Globals[wb.callerIdx]` or `vm.stack[callerSlot]`).
     - Correctly manage reference tracking: increment ref count on new assignment and decrement on overwritten slot without dropping native object pointer mappings.


---

## Verification & Validation Protocol
1. **Compilation Check:** Compile `./axonasp-http.exe` and `./axonasp-cli.exe` via `go build`.
2. **Go Unit Tests:** Run `go test ./axonvm/...` to verify stack integrity, opcode synchronization, and zero regressions.
3. **Regression Guard:** Verify that JScript AST parsing remains strictly isolated in `./jscript/` and zero-allocation stack discipline in VBScript VM is preserved.
