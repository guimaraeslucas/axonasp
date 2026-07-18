# ECMAScript Modules (import and export)

## Syntax

```javascript
import "./side-effects.js";
import { add, mul as multiply } from "./math.js";
import square, { PI } from "./math.js"; // Default and named
import * as ns from "./utils.js"; // Namespace import

export var version = "1.0";
export function sum(a, b) { return a + b; }
export default function(x) { return x * x; } // Default export
export { sum as add };
export { sum as addAlias } from "./math.js";
export * from "./other.js"; // Wildcard re-export
export * as ns from "./other.js"; // Namespace re-export
```

## Remarks

- `import` and `export` are supported for server-side JavaScript modules loaded from `.js` files.
- Module loading is **synchronous**. The VM resolves and executes imported modules in the same request execution flow.
- Module instances are stored per request in a request-local module registry. The same module path executes only once per request and subsequent imports reuse the same module environment.
- Compiled module bytecode uses the global script cache. This avoids recompilation when the source did not change.
- Circular dependencies are supported with partial initialization semantics.
- Standard ASP objects (`Response`, `Request`, `Session`, `Application`, `Server`) are automatically available inside modules.
- **ReferenceError:** The VM throws a `ReferenceError` if a requested named export is missing from the source module.
- **Global AST Cache:** Modules are read and compiled into AST/Bytecode ONCE globally and shared across all requests.
- **Request-Local Registry:** Each request has its own isolated module execution state. Top-level variables in a module are NOT shared between different users or subsequent requests.
- **Singleton per Request:** A module is executed only once within a single request, even if imported multiple times.
- **VM Reset:** Module instances are automatically cleared at the end of each request to prevent memory leaks and state contamination.
- **Module Resolution:** Imports are resolved relative to the current file path. Absolute paths and standard ASP virtual paths are also supported.

## Code Example

```javascript
<script runat="server" language="JScript">
// Assume 'config.js' exists with: export const version = "2.0";
import { version } from './config.js';
Response.Write("Application Version: " + version);
</script>
```
