# System Context & Role
Act as an Expert GoLang Developer, Systems Architect, Vue.js Frontend Engineer, and Server-Side Javascript (Classic ASP JScript/ECMA5 with ES6 polyfills) Engineer. Your task is to build a modern Reactive Component Framework (similar to ASP.NET WebForms or Laravel Livewire) built directly into the AxonASP Virtual Machine. 

We will call this native library **G3AxonLive**. It must allow developers to build stateful ASP pages where components update asynchronously without full page reloads, utilizing `fetch` and `WebSocket` support while keeping all business logic strictly inside ASP on the server. You must start and follow the phases as per user request and check completed when done.

# 🛑 Architectural Directives (CRITICAL)
1. **Native VM Integration & Procedural Control:** AxonLive must be built as a native Go library inside the `axonvm` package. The Go backend MUST absorb all request routing, JSON parsing, and JSON patch generation. Avoid complex VBScript/JScript wrapper classes (No `axonlive.asp` wrapper). The Go struct must act as the primary request controller.
2. **Performance & Memory:** Zero-allocation mindset. Do not cause Garbage Collector overloads in Go. Process HTTP streams directly into Go structs.
3. **No Reflection:** Strictly avoid Go `interface{}` and `reflect`. Use the existing `axonvm.Value` struct for all data passing.
4. **Language Constraint:** All code, comments, and documentation MUST be strictly in US English.
5. **Aesthetic (For the Builder):** The Vue.js builder UI must strictly follow the retro Microsoft MSDN Era (2003-2005) / Windows XP aesthetic. Rigid geometry (`border-radius: 0`), Tahoma/Verdana fonts, `#ECE9D8` backgrounds, and solid blue/gray borders.
6. **Safety & Security (Anti-Injection):** The framework and the Builder MUST prevent unauthorized code injection. Component IDs and Event Names mapped from the client JSON must be strictly validated against registered server-side components. Never use `eval()` or direct execution of client-provided payload strings in JS/VBScript. All shared state access must be synchronized (`sync.RWMutex`).
7. **Memory Management:** Implement robust background memory cleanup for orphaned sessions to prevent leaks.

# [x] 🛠️ Phase 1: Core Go State Management & Procedural Controller (`axonvm/lib_g3axonlive.go`)
Refactor the native Go struct `G3AXONLIVE` to act as the master controller for the page lifecycle.
* Implement standard VM signatures (`DispatchMethod`, `DispatchPropertyGet`).
* **New Methods & Properties:** Expose `InitPage()` (to parse the incoming X-G3AxonLive fetch request and extract JSON natively in Go), `IsAsyncRequest` (boolean), `EventComponentID`, `EventName`, and `EndAsyncResponse()` (which collects all updated component HTML, serializes the JSON patch response in Go, flushes to the HTTP writer, and cleanly halts ASP execution).
* Implement a timer mechanism that allows server-side code to trigger client-side events after a delay (e.g., `SetTimer(componentId, eventName, delay)`). Don't forget to handle timer cleanup and edge cases (e.g., if the session expires before the timer fires).
* Manage state safely with `sync.RWMutex` keyed by Session ID. Implement the background cleanup goroutine tied to `g3axonlive_active=true` in `axonasp.toml`. Remember to use the main global viper to avoid using more memory than necessary. 
* Remember that error messages must be saved in `axonvmerrorcodes.go` with a number and a descriptive message and called as the way the server handles errors in the endpoint, don't hardcode them.
* Don't let memory grow indefinitely. Follow the runtime limit of pages and components per session as defined in the config in the `global.default_script_timeout` obtained from viper.

# [x] 🛠️ Phase 2: The Communication Endpoint (`axonlive_handler.go`)
Refactor/Implement the Go HTTP handlers to process requests directly in the host server.
* Register endpoint `/g3al/` capable of handling async requests (if active in config/build tags).
* Accept POST requests containing JSON payloads.
* Draft the skeleton for a WebSocket handler upgrade on the exact same route to future-proof the architecture for continuous real-time connections.
* Remember that error messages must be saved in `axonvmerrorcodes.go` with a number and a descriptive message and called as the way the server handles errors in the endpoint, don't hardcode them.


# [x] 🛠️ Phase 3: The JavaScript Bridge (Fetch & DOM Swapping)
Refactor the client-side engine (The server-side wrapper is now obsolete).
* **Client-Side JS (`./www/axonlive/g3axonlive.js`):** Implement a vanilla JS engine. It intercepts UI events, prevents default actions, sends a `fetch` request to `/g3al/` with the header `X-G3AxonLive: true`, parses the returned JSON patch, and smartly swaps only the modified DOM elements using `outerHTML`.
* Ensure the client logic sanitizes output and gracefully handles network errors.
* The script must also be able to return errors on the console as it already does
* It must implement a retry mechanism with exponential backoff for transient network errors.
* It also must implement a timer and other helper functions, like redirect, Triggers, add attributes,  that can be called from the server-side to trigger client-side actions (for example, a `SetTimer(componentId, eventName, delay)` function that triggers an event after a delay). This is like the `__doPostBack` function in ASP.NET WebForms but with more capabilities and triggered from the server-side.

# [x] 🛠️ Phase 4: The Procedural Example Pages (`./www/axonlive/counter.asp` & `./www/axonlive/counter_js.asp`)
Write two complete procedural examples demonstrating the new Go-driven framework. Don't use classes. 
* **`counter.asp` (VBScript):** Demonstrate a top-down, procedural approach. Call `AxonLive.InitPage()`, check `If AxonLive.IsAsyncRequest`, handle the specific `AxonLive.EventComponentID`, update state, and call `AxonLive.EndAsyncResponse()`.
* **`counter_js.asp` (Server-Side Javascript):** Create the exact same counter logic using `<%@ Language="Javascript" %>`. Demonstrate how using JS functions and procedural flow interacts with the Go API cleanly. 
* Delete the file `axonlive.asp` as it is now obsolete and all logic is handled in Go and the client-side JS bridge.

# [x] 🛠️ Phase 5: Manual Update
Rewrite the manual section for G3AxonLive to reflect the new Procedural/Go-Controller architecture.
* Explicitly remove references to the old `AxonLive_Page` ASP class.
* Document the new Go-exposed methods (`InitPage`, `IsAsyncRequest`, `EventComponentID`, `EndAsyncResponse`).
* Add a dedicated section explaining how to use Server-Side Javascript with G3AxonLive, highlighting its advantages for modern developers.
* Explain the security model: how state is safely kept in Go memory and not tampered with by the client.
* Be specific, read the code to implement a helper adequated to the library. The user must have all the information on how to implement, features like timer, redirect, add attributes, and other helper functions that can be called from the server-side to trigger client-side actions must be explained in detail with examples.
* Include ASCII diagrams of the new architecture and flow.
* Include troubleshooting tips for common issues (e.g., "Why isn't my event firing?", "How do I debug the JSON patch response?", "What do the error codes mean?").
* Update the manual for G3AXONLIVE Methods, Overview, G3AxonLive Guides
* Give full examples in JavaScript, and in VBScript, and explain the differences and advantages of each.
* Update the menu to reflect the changes in G3AXONLIVE, DON'T CHANGE ANY OTHER PART OF THE MENU, JUST ADD THE NEW SECTIONS AND UPDATE THE EXISTING ONES RELATED TO G3AXONLIVE.

# [x] 🛠️ Phase 6: The Vue.js Visual Builder IDE
Architect a SPA in Vue.js that acts as a visual drag-and-drop builder for AxonLive pages in the folder (`/www/axonlive/builder/`). This builder will allow users to visually construct their ASP pages with AxonLive components and generate the corresponding code.
* **UI/UX:** Must strictly look like Visual Studio 2003 / Windows XP. No modern rounded UI frameworks.
* **Code Generation Pivot:** The Builder MUST generate **Server-Side Javascript** (JScript) instead of VBScript. JS is easier to manipulate, handles data structures better, and creates a cleaner output for the user.
* **Output:** As elements are dragged, generate:
  1. The procedural Server-Side JS logic mapping events to functions.
  2. The HTML required for rendering.
  3. A live preview pane that simulates the final rendered page.
  4. A JSON/tree representation of the component tree and event mappings for debugging.
  5. A "Copy to Clipboard" button that outputs the complete ASP page code (HTML + Server-Side JS) ready to be pasted into an `.asp` file.
  6. A "Download ASP File" button that allows users to download the generated code as an `.asp` file directly.
  7. Component properties panel where users can set attributes, styles, and event handlers with a form-based interface.
  8. Default event handlers for common components (e.g., buttons, inputs) that users can customize.
  9. Component library with pre-built elements (e.g., counters, forms(all types of inputs), labels, images, modals, tables, panels, placeholders, links, datalists) that users can drag and drop.
  10. A timer component that allows users to set server-side timers that trigger client-side events after a specified delay.

* **Security:** The generated code must utilize strict `switch/case` or mapping objects for event routing to guarantee that only known, developer-defined functions are triggered by client events. Prevent arbitrary code execution. Start elements with a unique prefix (e.g., `axl_`) to avoid ID collisions and ensure they are properly namespaced in the server-side logic.


# [ ] 🛠️ Phase 7: Builder Manual Update
Implement the manual for the Vue.js Visual Builder IDE.
* Detail how it generates Server-Side Javascript.
* Explain how users can safely inject their own business logic into the generated event handlers without breaking the builder's generated structure.
* Include screenshots/ASCII diagrams of the retro UI and step-by-step instructions.

# 📦 Output Requirements
Provide the refactored Go code for `lib_g3axonlive.go`, the updated HTTP endpoint snippet, the Vanilla JS client script, BOTH the VBScript and JScript procedural examples (`counter.asp` and `counter_js.asp`), and the Vue.js builder component structure. Ensure extreme backend efficiency and strict adherence to the visual aesthetic directives.
