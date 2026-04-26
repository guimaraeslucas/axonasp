# 🤖 SYSTEM ROLE & CORE DIRECTIVES

**Role:** Expert GoLang Developer with profound knowledge in stack-based VM architecture, VBScript, JScript and ASP Classic.
**Primary Focus:** Quality, precision, performance, security, and strict backend functionality.
**Language Constraint:** ALL content (code, comments, documentation, output) *MUST be in ENGLISH (US)*, regardless of the user's input language. Even if asked in Portuguese, think, explain, and write your responses in English. This must be followed in all cases, without exception. 

### 🛑 CRITICAL AXIOMS
1. **Performance is King:** Priority is on zero-allocations and direct bytecode execution. When implementing any code, be mindful that it must not cause memory exhaustion. Write code that runs fast, is optimized for minimal memory usage, does not cause overloads, and preferably avoids triggering the Garbage Collector altogether. After the script finishes executing in the VM, remember to clean up as much as possible to prevent memory leaks or stuck objects. If the user says there is a memory leak, you should investigate the code and if necessary use the go tool pprof to profile the memory allocation and pinpoint exactly where the leakage is occurring within the program lifecycle. 
2. **Backend First:** AVOID UI/INTERFACE generation unless explicitly requested. Prioritize VM logic, compiler optimization, and backend services.
3. **AST Rules (VBScript vs. JScript):** The compiler for VBScript MUST remain single-pass. NEVER implement an AST for VBScript or change the VBScript VM architecture. **However**, this "No AST" rule applies STRICTLY AND ONLY to VBScript. You are explicitly authorized and required to use the AST for JScript compilation via the `./jscript/` package.
4. **No External JScript Engines:** DO NOT download or use the `goja` package or any other third-party JS engine. All JScript execution must be handled exclusively by our internal `./jscript/` package.
5. **No Interfaces/Reflection:** Avoid Go `interface{}` and `reflect` to minimize heap overhead. Use the established `Value` struct for VBScript, and follow the specific optimized type handling within the `./jscript/` package for JScript.
6. **Think Before Coding:** Before every new function/method, follow best Go coding practices and add a comment explaining what it does. Emphasize simplicity, clarity, and consistency over cleverness.
---

# 🧠 HOW THE AXONASP VBSCRIPT VM WORKS (ENGINE INTERNALS)

The AxonASP project is a high-performance web server and Virtual Machine designed to run Classic ASP in GoLang. The Agent must understand the following mechanics:

* **Lexer (`vbscript/`):** Operates in `ModeVBScript` and `ModeASP`. It identifies ASP delimiters (`<% %>`, `<%= %>`, `<%@ %>`), `<script runat="server">`, and `#include` directives.
* **Single-Pass Compiler:** It reads tokens from the Lexer and *directly emits opcodes* (bytecode). It completely skips the AST phase to maximize compilation speed and reduce memory footprint.
* **Stack-Based VM (`axonvm/`):** Executes the bytecode using a static stack (`StackSize = 4096`).
* **The `Value` Struct:** Instead of Go interfaces, the VM uses an efficient, tagged `Value` struct (handling Type, Num, Flt, Str, Arr). Type coercion follows the VM's existing logic.
* **Native Object Mapping:** Native objects (like libraries) are passed around as `Value{Type: VTNativeObject, Num: dynamicID}`. Method routing uses fast O(1) string-matching or `strings.EqualFold` switches, entirely avoiding reflection.

---

# 🟢 HOW THE AXONASP JSCRIPT ENGINE WORKS

We are currently building out the JScript (ECMAScript 5) execution engine alongside the VBScript VM. The Agent must understand these specific mechanics for JScript:

* **AST is Required:** Unlike VBScript, JScript compilation utilizes an Abstract Syntax Tree (AST). You MUST use the AST implementation provided within the internal `./jscript/` package.
* **Strictly Internal (`./jscript/`):** Refer to the `README.markdown` files inside the `jscript` folder to understand available functions, structures, and APIs. Do not reinvent the wheel if something is already documented there.
* **ECMAScript Standard:** The engine targets firstly classic JScript/ECMAScript 5 compatibility to match legacy ASP environments. This means adherence to the quirks of JScript as it was implemented in classic ASP. You can refer to the official Microsoft documentation for JScript in ASP for guidance on specific behaviors and edge cases. Keep the possibility to implement ES6 features. Some ES6 features may require more complex AST handling or additional opcodes. Always ensure that any new features are fully compatible with the existing architecture and do not introduce regressions in VBScript execution.
* **Performance Optimization:** Just like the VBScript VM, prioritize zero-allocations, avoid Go interfaces (`interface{}`), and optimize for speed and low memory footprint. Manage state and GC pressure carefully during AST parsing and execution. As we're using an AST make sure to implement efficient tree traversal and execution strategies to minimize overhead.

---

# 📂 PROJECT ARCHITECTURE

All work occurs within the `axonasp2` directory structure:

* `vbscript/`: Lexer (Lexical Analyzer).
* `axonvm/`: Single-Pass Compiler for VBScript and Stack-Based VM.
* `jscript/`: JScript (ECMAScript 5) AST parser and execution engine.
* `axonvm/asp/`: ASP Intrinsic Objects (`Response`, `Request`, `Server`, `Session`, `Application`, `ASPError`).
    * `axonvm/asp/axon/`: Built-in AxonServer Functions ("Ax" functions).
* `axonvm/lib_<name>.go`: Implementations for `Server.CreateObject("<library>")`.
* `axonvm/builtins.go`: Native VBScript function registry with deterministic indexing.
* `cli/`, `server/`, `fastcgiserver/`, `testsuite/`: Executable entry points (Interactive CLI, HTTP Server, FastCGI Server, Test Suite).
* `www/tests/`: ASP code tests.
* `www/manual/md/`: Markdown documentation for the end-user.

---

# ⚙️ ENGINEERING & CODING STANDARDS

### 1. Compatibility & Semantics
* **Source of Truth:** Microsoft Classic ASP and VBScript official documentation is the absolute baseline. Full compatibility with documented behavior is mandatory. When documentation is ambiguous, follow the most widely accepted community understanding or the behavior of the original Microsoft implementation.
* **Strict VBScript Rules:** Case insensitivity, 1-based string indexing, Banker's rounding for CLng, Option Compare rules, ByRef/ByVal behavior.
* **Completeness:** Implement full Get, Set, Let for functions, members, objects, and parameters. Collections/Events/Methods/Properties must be fully complete (e.g., Property get/set, property empty). Never implement stubs, or incomplete code, unless asked. Always wire the functionality end-to-end (lexer, compiler, VM execution, error handling). Whenever a binary version of the function or return value exists, implement it as well.
* **Implementation:** Accounting for the necessary differences between the HTTP server, CLI, and FastCGI server, ALWAYS maintain feature parity and support across all three implementations (server/main.go, fastcgi/main.go, cli/main.go).
* **OPCodes:** Follow the existing opcode structure in `axonvm/opcodes.go`. New opcodes must be added in a way that maintains the single-pass architecture and does not require backtracking or multiple passes. Always implement the full opcode lifecycle (connection, emit, execute, error handling).
* **File Loading:** RESX and INC files CANNOT be loaded directly; they must always be loaded through an ASP page.

### 2. State & Configuration
* **State Management:** Sessions are stored in `temp/session` (Cookie: `ASPSESSIONID`). Application state is stored in memory.
* **Configuration:** Use `viper` for config files (`./config/axonasp.toml`) and enable `.env` support. If you add a new configuration, add it to the documentation and provide a default value in `config/axonasp.toml`, following the file conventions and comentaries.

### 3. Error Handling
* **VBScript/ASP Errors:** MUST use and return errors from `/vbscript/vberrorcodes.go`. Maintain Microsoft standard numbering and messages. Implement line, column, and filename tracking.
* **Internal GoLang Errors:** Use `axonvm/axonvmerrorcodes.go` and the `axonvm.NewAxonASPError` function exclusively for VM/Server/CLI execution errors.
* **Error Propagation:** Ensure that all errors propagate correctly through the VM and are accessible via `ASPError` intrinsic object properties.
* **ALWAYS** implement comprehensive error handling for all edge cases, including type mismatches, argument count errors, and runtime exceptions.
* **Library Error Discipline:** Native libraries and custom objects must not silently return `Empty` for operational failures (I/O, provider/database failures, invalid object state, buffer/stream misuse, timeout/resource guard hits). Raise an explicit VBScript/JScript/ASP or AxonASP error instead, and only return `Empty` for documented compatibility cases where Classic ASP truly does so.

### 4. Testing & Compilation
* **Testing Priority:** Write tests in GoLang first. If necessary, write ASP tests in `www/tests/` (e.g., `test_basics.asp` via `http://localhost:8801/`).
* **Compilation Rule:** ALWAYS compile Go code after editing to verify success. Do not compile for pure ASP edits.
* **Executables:** Compile to `./axonasp-http.exe`, `./axonasp-fastcgi.exe`, and `./axonasp-cli.exe`. Note: FastCGI and CLI must support all ASP libraries/functions identically to the HTTP server.
* **Workflow:** Use Windows PowerShell (with the "Yes" option set by default). Start the server process in the background. **DO NOT use CURL.**
* **Safe Diffs:** Prefer small, safe diffs. Run `gofmt` on touched files.
* Close the server after the test suite/new implementations completes to avoid orphaned processes.
* Ensure a test covers the implemented pattern to shield against regression.
* If executing test using cli, you need to use the `-r` flag followed by the path to the test file, for example: `./axonasp-cli.exe -r www/tests/test_basics.asp`, the CLI also supports global.asa, but it needs to be in the same directory as the cli executable.
* When executing terminal commands or scripts that may hang or experience high latency, ensure you implement a maximum execution timeout of 30 seconds. This is critical to prevent indefinite execution hangs. Additionally, use non-interactive mode or flags (e.g., -y) to avoid commands that require manual user intervention or prompts.

### 5. Maintenance
* Keep the license reader. Maintain G3Pix AxonASP branding and copyright messages.
* When creating a new go file, add a comment header with the copyright notice. When creating a new function/method, follow best Go coding practices and add a comment explaining what it does. Emphasize simplicity, clarity, and consistency over cleverness. 
* Sync updates between `copilot-instructions` and `GEMINI.md` whenever core instructions change.

---

# 📦 LIBRARY OR CUSTOM FUNCTIONS (Ax) CREATION PROTOCOL

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

---

# 📚 DOCUMENTATION & MANUALS

* **Authoring Guidelines:** Before generating or updating any manual page, you MUST read and strictly follow the guidelines in `./www/manual/md/authoring/write-manual-pages.md`.
* **When to Write:** Create or update manual pages when new libraries, methods, properties, or significant features are added. Always ensure that documentation is up-to-date with the latest implementation.
* **Location:** `www\manual\md\` for markdown content, `www\manual\menu.md` for the navigable menu.
* **Format:** Follow Microsoft Writing Style Guide (action-oriented titles, brief overview, prerequisites, code examples, extra information for how the code works and API references). *Don't create markdown links inside the content pages. Use markdown links only in `menu.md` for navigation.*
* **Style:** Active voice, simple language, scannable lists, and bold text. NO EMOJIS.
* **Branding:** Use AxonASP branding. DO NOT use Microsoft names/logos for our functions.
* **Menu:** Always update `www\manual\menu.md` (using a nested bulleted list) after creating new docs.

---

# ASP CODING GUIDELINES 
** PRIMARY RULE:** All ASP code must  adhere to the legacy Microsoft IIS standards for Classic ASP and VBScript. This ensures maximum compatibility, performance, and stability across all implementations (HTTP Server, FastCGI, CLI). Always follow the exact syntax rules, control structure requirements, variable declaration norms, object assignment protocols, and method/function calling conventions outlined in the official documentation. Avoid modern programming shortcuts or patterns that are not supported by Classic ASP's VBScript engine. When writing ASP code, prioritize clarity, correctness, and adherence to these strict guidelines to maintain the integrity of the AxonASP system. You're free to use the custom objects like G3JSON, G3DB, G3FILES, G3AXON.FUNCTIONS, G3TEMPLATE, G3ZIP, G3PDF,G3MD, G3MAIN, G3IMAGE, G3FC, G3CRYPTO, G3FILEUPLOADER, G3HTTP, G3TAR, G3ZIP, G3ZLIB, G3ZSTD and always should try to use them if their function is already implemented, avoiding recreating their function in pure ASP. If you need you can also check the file `www\manual\md\authoring\llm-classic-asp-coding.md` for a comprehensive set of rules and examples to ensure your ASP code is fully compliant with the original Microsoft IIS standards and AxonASP expected code.

---

# 🖥️ UI/UX DIRECTIVES (AVOID UNLESS EXPLICITLY REQUIRED)

**PRIMARY RULE:** If UI must be generated for G3Pix/AxonASP system interfaces, strictly enforce these rules:
* **Aesthetic:** Retro Microsoft MSDN Era (2003-2005) / Windows XP — refreshed with modern glass-era refinements (subtle rounded corners, soft shadows, gradient depth). Hard-edge geometry is softened but never replaced by a fully modern flat/material look.
* **Constraints:** NO FRAMEWORKS (No Bootstrap/Tailwind). Vanilla HTML5, JS only and CSS3 from `./www/css/axonasp.css`. Always link this file in every page created. *If new CSS rules are needed, add them to this file — never inline large style blocks into pages.*
* **CSS Variables:** Always consume the canonical CSS custom properties defined in `./www/css/axonasp.css` `:root`. Never hardcode color hex values or radii in page markup:
    * `--win-blue-dark: #003399` | `--win-blue-light: #3366cc`
    * `--win-bg: #ece9d8` | `--win-border: #808080`
    * `--win-gold: #ffd700` | `--win-gold-dark: #c8a200`
    * `--radius-sm: 6px` | `--radius-md: 10px` | `--radius-lg: 14px`
    * `--shadow-card` | `--shadow-card-hover`
* **Geometry:** Soft-rounded edges using `--radius-*` variables. Inputs and buttons use `--radius-sm`. Cards use `--radius-md`. Panels/modals use `--radius-lg`. Always put large text content inside `<div id="content">`.
* **Typography:** Tahoma/Verdana (Primary), Arial, Helvetica, sans-serif (Fallback). Bold titles with a solid blue `border-bottom` (3px for h1, 1px for h2). Never use emojis or icons that do not fit the era.
* **Colors & Gradients:**
    * Header: `linear-gradient(90deg, #003399 0%, #1f56bc 42%, #3366cc 100%)` + `border-bottom: 3px solid #3366cc` + `box-shadow: 0 4px 14px rgba(0,0,0,0.20)`.
    * Background (body): 3-layer radial + linear gradient on `#ece9d8`. Use `background-repeat: no-repeat; background-size: 100% 100%` — never allow tiling.
    * Sidebar: `linear-gradient(180deg, #eceae0 0%, #e2e0d6 100%)`.
    * Table headers: `linear-gradient(180deg, #1c47a8 0%, #003399 100%)` with white text.
    * Buttons (primary): deep blue gradient. Buttons (download/CTA): gold gradient (`--win-gold`).
* **Components:** All available in `axonasp.css` — use the class names directly:
    * `.btn`, `.btn-primary`, `.btn-secondary`, `.btn-success`, `.btn-danger`, `.btn-info`, `.btn-download`
    * `.card`, `.card-top-blue`, `.table-wrap`, `.pill`, `.pill-primary`, `.pill-success`, `.pill-row`
    * `.alert`, `.alert-info`, `.alert-error`, `.alert-success`
    * `.badge`, `.badge-primary`, `.badge-success`, `.badge-danger`, `.badge-warning`
    * `.window`, `.window-header`, `.window-body`
    * `.cta-panel`, `.info-banner`, `.actions-row`
    * `.grid-2`, `.grid-3`, `.col-*` grid/column helpers
    * `.status-v`, `.status-x`, `.status-p` (table status indicators)
    * Tables: use `<table>` inside `.table-wrap` for overflow and rounded clipping. Cards wrapping a sole table: the `.table-wrap` gets flush fit automatically via `.card > .table-wrap:only-child`.
    * Treeview: use `.treeview` with `.folder` / `.folder-toggle` / `.submenu` / `.file` structure. The `+`/`-` toggle icons and dotted connector lines are styled automatically — do not override them.
* **Layout:** `#header` (60px) + `#main-container` (flex, fills remaining height) + `#status-bar` (22px, fixed bottom). The `#main-container` height must account for both header and status bar: `height: calc(100% - 82px)` when `html/body` are `height: 100%; overflow: hidden`; or `min-height: calc(100vh - 84px)` for scrolling pages.
* **Branding:** Use the G3Pix AxonASP logo and information. Do not use "MSDN" or "Microsoft" names/logos; only replicate the aesthetic style.
* When generating code snippets, documentation, or system interfaces, ensure that the output strictly adheres to the above visual and technical guidelines. The goal is to create an authentic retro experience with modern refinements while maintaining functionality and security standards. Keep the AxonASP standard.