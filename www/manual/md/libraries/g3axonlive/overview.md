# G3AxonLive Overview

G3AxonLive is a powerful, native library integrated directly into the AxonASP core, designed to build modern, stateful, and reactive web pages using Classic ASP (VBScript or JScript). It allows developers to create dynamic user interfaces where components can be updated asynchronously without requiring a full page reload, similar to modern frameworks like ASP.NET WebForms, Laravel Livewire, or Vue.js.

The key paradigm shift with G3AxonLive v2.0 is the move away from ASP-based wrapper classes to a high-performance, procedural model controlled directly by a native GoLang object. All complex logic—request parsing, state management, and DOM patch generation—is handled by the optimized Go backend, leaving you to focus on your application's business logic in a clean, procedural style.

## Key Features

*   **Native Performance:** All heavy lifting is done in compiled Go code, ensuring minimal overhead and maximum speed. State is managed securely in Go's memory space, protected from client-side tampering.
*   **Procedural Simplicity:** No complex classes or inheritance required. You interact with a clean, procedural API directly within your ASP pages.
*   **Dual Language Support:** Write your backend logic in either **VBScript** or **Server-Side JScript (ECMAScript 5)**. JScript is highly recommended for its superior data handling and modern syntax.
*   **Asynchronous DOM Updates:** G3AxonLive uses the browser's `fetch` API to send small JSON payloads to the server. The server responds with highly efficient DOM patches, which the client-side engine uses to update only the modified components.
*   **Server-Driven Actions:** The server can instruct the client to perform actions like redirecting to a new page, setting a timer to trigger a future event, or modifying a component's attributes, giving you full control over the user experience from your backend code.
*   **Zero Dependencies:** The client-side engine is a lean, dependency-free vanilla JavaScript file (`g3axonlive.js`), ensuring broad compatibility and fast load times.

G3AxonLive empowers you to build rich, interactive applications with the simplicity and familiarity of Classic ASP, supercharged by a modern, high-performance Go backend.
