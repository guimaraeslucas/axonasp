When implementing custom functions for AxonASP (Ax), adhere to the following guidelines:

1.  **File Placement:** Create `axonvm/lib_<name>.go` or in `axonvm/asp/axon` for "Ax" (Custom) functions.
2.  **Implementation:** Define a concrete Go struct.
3.  **Strictly No Reflection:** Implement two strongly-typed, switch-based dispatch methods:
    * `DispatchMethod(methodName string, args []axonvm.Value) axonvm.Value`
    * `DispatchPropertyGet(propertyName string) axonvm.Value`
4.  **Type Safety:** Only use `axonvm.Value` and its constructors (`NewString`, `NewInteger`, etc.).
5.  **VM Integration (`axonvm/vm.go`):**
    * Update the VM struct with a map of active instances (e.g., `map[int64]*libraries.MyObject`).
    * Update `NewVM` to instantiate this map.
    * Update `dispatchNativeCall` (`Server.CreateObject`): Intercept PROGID, instantiate struct, assign dynamic ID, store in map, return `VTNativeObject`.
    * Update `dispatchNativeCall` and `dispatchMemberGet` switch blocks to route method/property calls to your dispatch functions based on the dynamic ID.
6.  **Error Handling:** Implement full error support for argument count/type failures, attaching filename/line/col mimicking ASP error reporting. You can find the ASP/VBScript errors in `vbscript/vberrorcodes.go`, the ASP/JScript error in `jscript/jscripterrorcodes.go`, and the internal/VM/AxonASP errors in `axonvm\axonvmerrorcodes.go` (for the custom functions, VM internal errors and custom libraries, you should implement an error number/description in this file so it is reusable and the user can have better debug info - never implement a hardcoded error/string, always implement in this file and then get the value from it). When a new error code is implemented, update the documentation with the new error and its meaning, and if it is an error that can be raised by the user code, in the file `www\manual\md\runtime\axonasp-error-codes.md`.
7. When implementing any code, be mindful that it must not cause memory exhaustion. Write code that runs fast, is optimized for minimal memory usage, does not cause overloads, and preferably avoids triggering the Garbage Collector altogether. After the script finishes executing in the VM, remember to clean up as much as possible to prevent memory leaks or stuck objects.