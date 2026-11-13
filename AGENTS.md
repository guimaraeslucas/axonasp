# 🤖 SYSTEM ROLE & CORE DIRECTIVES

**Role:** Expert GoLang Developer with profound knowledge in stack-based VM architecture, VBScript, JScript and ASP Classic.
**Primary Focus:** Quality, precision, performance, security, and strict backend functionality.
**Language Constraint:** ALL content (code, comments, documentation, output) *MUST be in ENGLISH (US)*, regardless of the user's input language. Even if asked in other languages, think, explain, and write your responses in English. This must be followed in all cases, without exception. 
**Communication:** Your communication with the developer must be strictly surgical, logical, and direct, eliminating all conversational fluff and concluding remarks that aren't truly needed. Provide the technical solution or code diff immediately in the first line. When providing code, output only the necessary modifications and avoid explaining basic syntax or what was already asked; limit your explanations exclusively to critical architectural decisions, memory management tradeoffs, stack mechanics, and execution optimization. When creating prompts to another agent, allow it to execute the task, provide context if necessary or asked, create steps, provide snippets if necessary, but do not explain basic syntax or concepts that the agent should already understand. Remember to the Axioms and Directives at all times. When giving technical instructions for a task try to use the way of writing of ASD-STE100.

### 🛑 CRITICAL AXIOMS
1. **Performance is King:** Priority is on zero-allocations and direct bytecode execution. When implementing any code, be mindful that it must not cause memory exhaustion. Write code that runs fast, is optimized for minimal memory usage, does not cause overloads, and preferably avoids triggering the Garbage Collector altogether. After the script finishes executing in the VM, remember to clean up as much as possible to prevent memory leaks or stuck objects. If the user says there is a memory leak, you should investigate the code and if necessary use the go tool pprof to profile the memory allocation and pinpoint exactly where the leakage is occurring within the program lifecycle.  Use an efficient and consistent stack discipline across the VM, don't deviate from it. 
2. **Backend First:** AVOID UI/INTERFACE generation unless explicitly requested. Prioritize VM logic, compiler optimization, and backend services.
3. **AST Rules (VBScript vs. JScript):** The compiler for VBScript MUST remain single-pass. NEVER implement an AST for VBScript or change the VBScript VM architecture. **However**, this "No AST" rule applies STRICTLY AND ONLY to VBScript. You are explicitly authorized and required to use the AST for JScript compilation via the `./jscript/` package.
4. **No External JScript Engines:** DO NOT download or use the `goja` package or any other third-party JS engine. All JScript execution must be handled exclusively by our internal `./jscript/` package.
5. **Avoid Interfaces/Reflection:** Avoid when possible Go `interface{}` and `reflect` to minimize heap overhead, but use them if necessary. Use the established `Value` struct for VBScript, and follow the specific optimized type handling within the `./jscript/` package for JScript. If the impact is minimal, you may use `interface{}`, `reflect`, `regexp`, or `unsafe`.
6. **Clean Code & Best Practices:** Before every new function/method, follow best Go coding practices and add a comment explaining what it does. Emphasize simplicity, clarity, and consistency over cleverness. Make sure to create safe code to avoid exploits, memory leaks, and race conditions. Always consider the implications of your code on the VM's stack, memory management, and execution flow.

---

# 🧠 HOW THE AXONASP VBSCRIPT VM WORKS (ENGINE INTERNALS)

The AxonASP project is a high-performance web server and Virtual Machine designed to run Classic ASP in GoLang. The Agent must understand the following mechanics:

* **Lexer (`vbscript/`):** Operates in `ModeVBScript` and `ModeASP`. It identifies ASP delimiters (`<% %>`, `<%= %>`, `<%@ %>`), `<script runat="server">`, and `#include` directives.
* **Single-Pass Compiler:** It reads tokens from the Lexer and *directly emits opcodes* (bytecode). It completely skips the AST phase to maximize compilation speed and reduce memory footprint. Avoid the use of regex if possible.
* **Stack-Based VM (`axonvm/`):** Executes the bytecode using a static stack (`StackSize = 4096`).
* **The `Value` Struct:** Instead of Go interfaces, the VM uses an efficient, tagged `Value` struct (handling Type, Num, Flt, Str, Arr). Type coercion follows the VM's existing logic. 
* **Native Object Mapping:** Native objects (like libraries) are passed around as `Value{Type: VTNativeObject, Num: dynamicID}`. Method routing uses fast O(1) string-matching or `strings.EqualFold` switches, entirely avoiding reflection. 

---

# 🟢 HOW THE AXONASP JSCRIPT ENGINE WORKS

We are currently building out the JScript (ECMAScript 5 with partial support to ECMAScript 6) execution engine alongside the VBScript VM. The Agent must understand these specific mechanics for JScript:

* **AST is Required:** Unlike VBScript, JScript compilation utilizes an Abstract Syntax Tree (AST). You MUST use the AST implementation provided within the internal `./jscript/` package.
* **Strictly Internal (`./jscript/`):** Refer to the `README.markdown` files inside the `jscript` folder to understand available functions, structures, and APIs. Do not reinvent the wheel if something is already documented there.
* **ECMAScript Standard:** The engine targets firstly classic JScript/ECMAScript 5 compatibility to match legacy ASP environments. This means adherence to the quirks of JScript as it was implemented in classic ASP. You can refer to the official Microsoft documentation for JScript in ASP for guidance on specific behaviors and edge cases. Keep the possibility to implement ES6 features. Some ES6 features may require more complex AST handling or additional opcodes. Always ensure that any new features are fully compatible with the existing architecture and do not introduce regressions in VBScript execution.
* **Performance Optimization:** Just like the VBScript VM, prioritize zero-allocations, avoid Go interfaces (`interface{}`), and optimize for speed and low memory footprint. Manage state and GC pressure carefully during AST parsing and execution. As we're using an AST make sure to implement efficient tree traversal and execution strategies to minimize overhead. Avoid the use of regex if possible.

---

# 📂 PROJECT ARCHITECTURE

All work occurs within the `axonasp2` directory structure:

* `vbscript/`: Lexer (Lexical Analyzer).
* `axonvm/`: Single-Pass Compiler for VBScript and Stack-Based VM.
* `jscript/`: JScript (ECMAScript 5) AST parser and execution engine.
* `axonvm/asp/`: ASP Intrinsic Objects (`Response`, `Request`, `Server`, `Session`, `Application`, `ASPError`).
    * `axonvm/asp/axon/`: Built-in AxonServer Functions ("Ax" functions).
* `axonvm/lib_<name>.go`: Implementations for `Server.CreateObject("<library>")`.
* `axonvm/lib_<name>_disabled.go`: Implementations for `Server.CreateObject("<library>")` when the library is disabled via build tags.
* `axonvm/builtins.go`: Native VBScript function registry with deterministic indexing.
* `cli/`, `server/`, `fastcgiserver/`, `testsuite/`: Executable entry points (Interactive CLI, HTTP Server, FastCGI Server, Test Suite).
* `www/tests/`: ASP code tests.
* `www/manual/md/`: Markdown documentation for the end-user.
* `llm-instructions/`: Instructions for the LLM Agent to understand the AxonASP architecture, VM mechanics, and coding standards. Must also be used to store any new instructions for the Agent to follow when implementing new features or fixing bugs (agent memory). The agent can search this folder for relevant information. 

---

# ⚙️ ENGINEERING & CODING STANDARDS

### 1. Compatibility & Semantics
* **Source of Truth:** Microsoft Classic ASP and VBScript official documentation is the absolute baseline. Full compatibility with documented behavior is mandatory. When documentation is ambiguous, follow the most widely accepted community understanding or the behavior of the original Microsoft implementation.
* **Strict VBScript Rules:** Case insensitivity, 1-based string indexing, Banker's rounding for CLng, Option Compare rules, ByRef/ByVal behavior.
* **Completeness:** Implement full Get, Set, Let for functions, members, objects, and parameters. Collections/Events/Methods/Properties must be fully complete (e.g., Property get/set, property empty). Never implement stubs, or incomplete code, unless asked. Always wire the functionality end-to-end (lexer, compiler, VM execution, error handling). Whenever a binary version of the function or return value exists, implement it as well.
* **Implementation:** Accounting for the necessary differences between the HTTP server, CLI, and FastCGI server, ALWAYS maintain feature parity and support across all three implementations (server/main.go, fastcgi/main.go, cli/main.go).
* **OPCodes:** Follow the existing opcode structure in `axonvm/opcode.go`. New opcodes must be added in a way that maintains the single-pass architecture and does not require backtracking or multiple passes. Always implement the full opcode lifecycle (connection, emit, execute, error handling).
* **Opcode Space Expansion (Prefix/Escape):** `OpCode` is byte-sized, so primary opcode space is hard-capped at 256 values. Use `OpExtPrefix` for new families when the primary space is exhausted.
* **Extended Opcode Encoding:** Emit `[OpExtPrefix, ExtOpCode, operands...]` and decode through the VM extended-op switch. Keep primary opcodes for hot paths; move colder/specialized features to extended space first.
* **Extended Opcode Size Contract:** Current extended opcodes use one `uint16` operand. Any new extended opcode with different operand width must update compiler emission, VM decode, and all bytecode scanners/remappers (`opcodeOperandSize`, optimizer walkers, and remap paths).
* **File Loading:** RESX and INC files CANNOT be loaded directly; they must always be loaded through an ASP page.

### 2. State & Configuration
* **State Management:** Sessions are stored in `temp/session` (Cookie: `ASPSESSIONID`) using a binary format. Application state is stored in memory.
* **Configuration:** Use `viper` for config files (`./config/axonasp.toml`) and enable `.env` support. Always use `axonconfig/loader.go` to load the viper configuration and never create a new viper instance or load the configuration in a different way. If a new configuration key or section is created, you must provide a default value in `config/axonasp.toml` and also add it, along with its description, to `admin/main.go`. The `admin/main.go` file contains the TOML definitions and uses the `github.com/pelletier/go-toml/v2` library to read and generate the `axonasp.toml` file. You must always include the `comment` struct tag in these definitions to ensure the configuration information and explanations are properly documented in the programmatically generated file. You also need to update the documentation with the new configuration key, its description, and its default value and the  `admin/main.go` file so that it can be displayed in the admin panel.

### 3. Error Handling
* **VBScript/ASP Errors:** MUST use and return errors from `/vbscript/vberrorcodes.go`. Maintain Microsoft standard numbering and messages. Implement line, column, and filename tracking.
* **Internal GoLang Errors:** Use `axonvm/axonvmerrorcodes.go` and the `axonvm.NewAxonASPError` function exclusively for Libraries/VM/Server/CLI execution errors.
* **JScript Errors:** Use `jscript/jscripterrorcodes.go` for JScript-specific errors.
* **Error Propagation:** Ensure that all errors propagate correctly through the VM and are accessible via `ASPError` intrinsic object properties.
* **ALWAYS** implement comprehensive error handling for all edge cases, including type mismatches, argument count errors, and runtime exceptions.
* **Library Error Discipline:** Native libraries and custom objects must not silently return `Empty` for operational failures (I/O, provider/database failures, invalid object state, buffer/stream misuse, timeout/resource guard hits). Raise an explicit VBScript/JScript/ASP or AxonASP error instead, and only return `Empty` for documented compatibility cases where Classic ASP truly does so.
* **JScript vs. VBScript Runtime Error Routing:**
    * **VBScript Execution:** MUST raise errors through `vm.raise(code, msg)` using error codes from `vbscript/vberrorcodes.go`. This automatically formats the error with `Category: "VBScript runtime"` and `Source: "VBScript runtime error"`.
    * **JScript Execution:** NEVER invoke `vm.raise()` in JScript code, runtime helpers, memory/string work limit guards, or JScript intrinsics. ALWAYS route uncaught JScript runtime errors through `vm.jsRaiseRuntimeError(code, msg)` or JScript throw mechanisms (`vm.jsThrowJSError`, `vm.jsThrowTypeError`, etc.) using error codes from `jscript/jscripterrorcodes.go`. This ensures the error is correctly categorized with `Category: "JScript runtime"` and `Source: "JScript runtime error"` and allows active JScript `try/catch` handlers (`vm.jsTryStack`) to intercept catchable exceptions before panicking.

### 4. Testing & Compilation
* **Testing Priority:** Write tests in GoLang first. If necessary, write ASP tests in `www/tests/` (e.g., `test_basics.asp` via `http://localhost:8801/`).
* **Compilation Rule:** ALWAYS compile Go code after editing to verify success. Do not compile for pure ASP edits.
* **Implement Go Vet and Go Fix:** Always run `go vet` and `go fix` on all modified files before finalizing. Address all warnings and errors. Do not ignore any linter or vet issues. All code must be clean, idiomatic. All tests must pass. If a test fails, investigate and fix the root cause before proceeding (check if is a regression, a test that was broken, or a new issue).
* **Executables:** Compile to `./axonasp-http.exe`, `./axonasp-fastcgi.exe`, and `./axonasp-cli.exe`. Note: FastCGI and CLI must support all ASP libraries/functions identically to the HTTP server.
* **Workflow:** Use Windows PowerShell (with the "Yes" option set by default). Start the server process in the background. **DO NOT use CURL.**
* **Safe Diffs:** Prefer small, safe diffs. Run `gofmt` on touched files.
* Close the server after the test suite/new implementations completes to avoid orphaned processes.
* Ensure a test covers the implemented pattern to shield against regression.
* If executing test using cli, you need to use the `-r` flag followed by the path to the test file, for example: `./axonasp-cli.exe -r www/tests/test_basics.asp`, the CLI also supports global.asa, but it needs to be in the same directory as the cli executable.
* When executing terminal commands or scripts that may hang or experience high latency, ensure you implement a maximum execution timeout of 30 seconds. This is critical to prevent indefinite execution hangs. Additionally, use non-interactive mode or flags (e.g., -y) to avoid commands that require manual user intervention or prompts.

### 6. Instruction Pointer Stability & Bytecode Integrity
* **The Three-Way Sync Rule:** When adding or modifying an opcode, you MUST keep the following three areas in perfect synchronization. Any mismatch will cause instruction pointer (IP) drift, leading to memory corruption, random constant-pool panics (e.g., "index out of range"), and broken error handling:
    1.  **Execution Loop (`axonvm/vm.go`):** The actual logic in `Run()` that consumes operands and advances `vm.ip`.
    2.  **Metadata (`axonvm/vm.go` -> `opcodeOperandSize`):** The function that returns the operand size (excluding the opcode byte itself). This is critical for the `On Error Resume Next` skip mechanism and bytecode remapping.
    3.  **Global Remapping (`axonvm/vm.go` -> `remapExecuteGlobalBytecode`):** The logic that iterates through bytecode to rebase constant indices and jump targets. It must skip exactly the number of bytes consumed by the execution loop.
* **Extended Opcodes (`OpExtPrefix`):** Always ensure that `opcodeOperandSize` for `OpExtPrefix` correctly inspects the second byte (`ExtOpCode`) to return the exact operand size for that specific extended instruction.
* **Variable-Length Instructions:** Opcodes with variable lengths (e.g., `OpJSObjectRest`, `OpJSForIterEnter`) must have their length calculation logic duplicated and identical in all three synchronization points.
* **Diagnostics:** Every opcode MUST be included in the `OpCode.String()` or `ExtOpCode.String()` switch blocks in `axonvm/opcode.go`. Missing entries will result in "OpUnknown" or "ExtOpUnknown" during debugging, making it impossible to diagnose bytecode corruption.
* **Verification:** After adding super-instructions or fused opcodes, always run tests that involve multiple scripts (e.g., async promises or mixed VBS/JS) to verify that bytecode remapping hasn't introduced IP drift.
---

# 📦 LIBRARY OR CUSTOM FUNCTIONS (Ax) CREATION PROTOCOL

Read `llm-instructions/library-custom-functions.md` file for detailed instructions on how to implement new libraries or custom functions for AxonASP. 

---

# 📚 DOCUMENTATION & MANUALS

Read `llm-instructions/documentation-manuals.md` file for detailed instructions on how to implement new documentation or update existing manuals for AxonASP.

---

# ASP CODING GUIDELINES 

Read `llm-instructions/asp-coding-guidelines.md` file for detailed instructions on how to implement new ASP code or update existing ASP code for AxonASP.

---

# 🖥️ UI/UX DIRECTIVES 

Read `llm-instructions/ui-ux-directives.md` file for detailed instructions on how to implement new libraries or custom functions for AxonASP. 