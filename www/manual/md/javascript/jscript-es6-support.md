# Use ES6 Features and beyond in JavaScript Scripts

## Overview

AxonASP's JavaScript engine supports a wide range of modern ECMAScript features, including ES6 (ES2015) additions and subsequent standards up to ES2024. This page provides a comprehensive index of all supported modern capabilities. Each feature has a dedicated subpage with full syntax, remarks, and code examples.

All ES6 features described here are available in `<script runat="server" language="JavaScript">` blocks or by setting the `<%@Language="JavaScript"%>` header.

---

## 1. Proxies

Create proxy objects that intercept and redefine fundamental operations on target objects. Supports all standard traps including `get`, `set`, `has`, `deleteProperty`, `apply`, `construct`, and `ownKeys`. Strict invariant enforcement follows ECMAScript section 10.5 rules.

Refer to the dedicated Proxies page for syntax, complete trap documentation, and code examples.

---

## 2. Reflect API

A built-in object providing methods for interceptable JavaScript operations. Methods mirror proxy handler traps and return success booleans instead of throwing exceptions. Includes `Reflect.get`, `Reflect.set`, `Reflect.has`, `Reflect.deleteProperty`, `Reflect.ownKeys`, `Reflect.apply`, and `Reflect.construct`.

Refer to the dedicated Reflect API page for syntax and code examples.

---

## 3. ECMAScript Modules (import and export)

Supports server-side JavaScript modules with `import` and `export` statements. Module loading is synchronous. Modules are compiled once globally and cached, with per-request isolated execution state. Circular dependencies with partial initialization semantics are supported.

Refer to the dedicated ECMAScript Modules page for syntax and code examples.

---

## 4. Weak Collections (WeakMap and WeakSet)

Collections where keys are held weakly, preventing memory leaks when objects are used as temporary keys. AxonASP uses an efficient inverted storage pattern for automatic data reclamation. Supports Objects, Functions, and unique Symbols as keys.

Refer to the dedicated Weak Collections page for syntax and code examples.

---

## 5. Weak References (WeakRef and FinalizationRegistry)

Provides weak reference semantics and cleanup callbacks via `WeakRef.deref()` and `FinalizationRegistry`. Fully implemented for JavaScript library compatibility. In the AxonASP short-lived request context, garbage collection callbacks are not triggered during standard execution.

Refer to the dedicated Weak References page for syntax and code examples.

---

## 6. Block-Scoped Declarations (let and const)

Block-level scoping with `let` and `const`. Temporal Dead Zone (TDZ) enforcement prevents access before declaration. `const` bindings are immutable; reassignment throws a `TypeError`.

Refer to the dedicated Block-Scoped Declarations page for syntax and code examples.

---

## 7. Explicit Resource Management (using)

Block-scoped resource management using `using` and `async using` declarations. Cleanup uses `Symbol.dispose` and `Symbol.asyncDispose`. Resources are disposed in reverse declaration order on scope exit and during exception unwinding.

Refer to the dedicated Explicit Resource Management page for syntax and code examples.

---

## 8. Full Unicode Support

Supports ES6 Unicode code point escapes (`\u{...}`) for characters up to `0x10FFFF`. The RegExp `u` flag enables full code point matching and Unicode escape sequences in patterns.

Refer to the dedicated Full Unicode Support page for syntax and code examples.

---

## 9. Modern Regular Expressions (RegExp)

PCRE-compatible engine supporting ES6+ features: named capture groups, lookahead and lookbehind assertions, sticky flag (`y`), and the `flags` property.

Refer to the dedicated Modern Regular Expressions page for syntax and code examples.

---

## 10. Template Literals

String literals using backticks (`` ` ``) with `${expression}` interpolation. Preserve literal newlines and support multiple embedded expressions. Tagged template literals are not supported.

Refer to the dedicated Template Literals page for syntax and code examples.

---

## 11. Arrow Functions

Concise function syntax with lexical `this` binding. Single-parameter parentheses are optional. Cannot be used as constructors. Arrow functions do not have their own `arguments` object.

Refer to the dedicated Arrow Functions page for syntax and code examples.

---

## 12. Default Parameter Values

Native default parameter syntax allows functions to specify default values for parameters when they are `undefined`. The classic guard pattern remains supported for compatibility.

Refer to the dedicated Default Parameter Values page for syntax and code examples.

---

## 13. Tail Call Optimization (TCO)

Optimizes tail-position calls by reusing the active function frame. Enforces a 10,000 instruction limit to prevent memory exhaustion. Disabled inside `try`, `catch`, or `finally` blocks. Use with caution.

Refer to the dedicated Tail Call Optimization page for syntax, remarks, and code examples.

---

## 14. Rest Parameters

Collects remaining function arguments into a standard array. Must be the last parameter in the function signature. Only one rest parameter is allowed per function.

Refer to the dedicated Rest Parameters page for syntax and code examples.

---

## 15. Object Literal Property Shorthand

Allows omitting the value when the variable name matches the property name. Method shorthand syntax is also supported.

Refer to the dedicated Object Literal Property Shorthand page for syntax and code examples.

---

## 16. Spread in Array Literals

Expands array-like values into individual elements within a new array. Left-to-right evaluation order. `null` and `undefined` spread sources raise a `TypeError`.

Refer to the dedicated Spread in Array Literals page for syntax and code examples.

---

## 17. Object Static Utilities

Modern `Object` static methods: `Object.assign`, `Object.keys`, `Object.values`, `Object.entries`, `Object.fromEntries`, `Object.is`, `Object.setPrototypeOf`, and `Object.getOwnPropertySymbols`.

Refer to the dedicated Object Static Utilities page for syntax and code examples.

---

## 18. Property Reflection Helpers

Reflection support via `Object.getOwnPropertyDescriptor` and `Object.getOwnPropertyDescriptors`. Returns property descriptors with `value`, `writable`, `enumerable`, and `configurable` fields.

Refer to the dedicated Property Reflection Helpers page for syntax and code examples.

---

## 19. Array Search Utilities

ES6 array search methods: `Array.prototype.find` returns the first matching element, and `Array.prototype.findIndex` returns the index of the first match. Both accept a callback and optional `thisArg`.

Refer to the dedicated Array Search Utilities page for syntax and code examples.

---

## 20. Array Construction Utilities

`Array.from` converts array-like or iterable objects into standard arrays with an optional mapping function. `Array.of` creates arrays from its arguments, avoiding the `new Array(n)` single-element pitfall.

Refer to the dedicated Array Construction Utilities page for syntax and code examples.

---

## 21. Array In-place Operations

Comprehensive array methods including `fill`, `copyWithin`, `keys`, `entries`, `at`, `flat`, `flatMap`, and immutable methods `toSorted`, `toReversed`, and `toSpliced` that return new arrays without mutating the original.

Refer to the dedicated Array In-place Operations page for syntax and code examples.

---

## 22. ES6 String Methods

Modern string utilities: `includes`, `startsWith`, `endsWith`, `repeat`, `at`, `codePointAt`, `normalize`, `padStart`, and `padEnd`. Regexp-aware `includes` throws `TypeError` for RegExp arguments.

Refer to the dedicated ES6 String Methods page for syntax and code examples.

---

## 23. ES6 Number Static Methods

`Number.isInteger`, `Number.isNaN`, `Number.isFinite`, `Number.isSafeInteger`, `Number.parseInt`, and `Number.parseFloat`. All methods do not coerce non-number values. Includes read-only constants like `Number.MAX_SAFE_INTEGER` and `Number.EPSILON`.

Refer to the dedicated ES6 Number Static Methods page for syntax and code examples.

---

## 24. Binary and Octal Numeric Literals

Binary literals use the `0b` or `0B` prefix. Octal literals use the `0o` or `0O` prefix. Both parse to standard numeric values.

Refer to the dedicated Binary and Octal Numeric Literals page for syntax and code examples.

---

## 25. Global URI Functions

`encodeURI`, `decodeURI`, `encodeURIComponent`, and `decodeURIComponent` for percent-encoding and decoding URI strings and components. `encodeURIComponent` escapes reserved characters like `=`, `&`, and `+`.

Refer to the dedicated Global URI Functions page for syntax and code examples.

---

## 26. Math Extensions

Extended `Math` methods: `Math.trunc`, `Math.sign`, `Math.cbrt`, plus hyperbolic, logarithmic, and miscellaneous functions including `Math.hypot`, `Math.imul`, and `Math.clz32`.

Refer to the dedicated Math Extensions page for syntax and code examples.

---

## 27. Symbol Primitive

Unique, immutable primitive values used as collision-safe object property keys. Each call to `Symbol()` returns a unique value. Symbol-keyed properties are hidden from `Object.keys`, `Object.values`, and `Object.entries`.

Refer to the dedicated Symbol Primitive page for syntax and code examples.

---

## 28. Symbol Primitive - Well-Known Symbols and Global Registry

Pre-defined well-known symbols including `Symbol.iterator`, `Symbol.toStringTag`, `Symbol.species`, `Symbol.hasInstance`, and `Symbol.toPrimitive`. Global symbol registry via `Symbol.for` and `Symbol.keyFor`.

Refer to the dedicated Symbol Well-Known page for syntax and code examples.

---

## 29. Iteration Protocol - for...of and Custom Iterables

The `for...of` loop iterates over iterable objects. Built-in iterables include Array, String, Set, and Map. Custom iterables require implementing the `[Symbol.iterator]` method returning an iterator with a `next()` method.

Refer to the dedicated Iteration Protocol page for syntax and code examples.

---

## 30. Binary Data - ArrayBuffer, SharedArrayBuffer and Typed Arrays

Raw binary data buffers via `ArrayBuffer` and `SharedArrayBuffer`. Typed array views include `Int8Array`, `Uint8Array`, `Uint8ClampedArray`, `Int16Array`, `Uint16Array`, `Int32Array`, `Uint32Array`, `Float32Array`, `Float64Array`, `BigInt64Array`, and `BigUint64Array`. Byte-level control via `DataView` with explicit endianness.

Refer to the dedicated Binary Data page for syntax and code examples.

---

## 31. Set and Map Collections

`Set` stores unique values with `add`, `has`, `delete`, `clear`, and `size`. `Map` stores key/value pairs with insertion order preservation and `set`, `get`, `has`, `delete`, `clear`, and `size`.

Refer to the dedicated Set and Map Collections page for syntax and code examples.

---

## 32. Computed Property Names

Use square bracket expressions inside object literals to compute property names at runtime. Any valid JScript expression can be used. Computed keys can be mixed with static keys and shorthand properties.

Refer to the dedicated Computed Property Names page for syntax and code examples.

---

## 33. Internationalization API (Intl)

Locale-sensitive formatting and comparison: `Intl.DateTimeFormat`, `Intl.NumberFormat`, `Intl.Collator`, `Intl.PluralRules`, and `Intl.RelativeTimeFormat`. Uses AxonASP locale profiles with fallback to `en-US`.

Refer to the dedicated Internationalization API page for syntax and code examples.

---

## 34. Destructuring Assignment

Unpack values from arrays (or any iterable) and properties from objects into distinct variables. Supports default values, rest patterns, elision, nesting, and computed property names.

Refer to the dedicated Destructuring Assignment page for syntax and code examples.

---

## 35. ES6 Classes

Modern class syntax with `constructor`, instance methods, static methods, getters/setters, private fields (`#property`), and inheritance via `extends` and `super`. Strict mode is implicit within all class bodies.

Refer to the dedicated ES6 Classes page for syntax and code examples.

---

## 36. Optional Chaining (?.)

Safely access deeply nested properties without explicit null checks. Short-circuits to `undefined` when the operand before `?.` is `null` or `undefined`. Works for property access, bracket access, and function calls.

Refer to the dedicated Optional Chaining page for syntax and code examples.

---

## 37. new.target meta-property

Detect whether a function or constructor was called with the `new` operator. Returns the constructor reference when called with `new`, and `undefined` for normal function calls.

Refer to the dedicated new.target page for syntax and code examples.

---

## 38. Nullish Coalescing (??)

Returns the right-hand side operand only when the left-hand side is `null` or `undefined`. Unlike `||`, it does not treat `0`, `""`, or `false` as values requiring a fallback.

Refer to the dedicated Nullish Coalescing page for syntax and code examples.

---

## 39. Logical Assignment (||=, &&=, ??=)

Short-circuit assignment operators: `||=` assigns when the target is falsy, `&&=` assigns when the target is truthy, and `??=` assigns when the target is nullish. Right-hand side is only evaluated when needed.

Refer to the dedicated Logical Assignment page for syntax and code examples.

---

## 40. Exponentiation Operator (**)

Returns the result of raising the first operand to the power of the second operand. Equivalent to `Math.pow()` with additional BigInt support. Includes the `**=` assignment variant.

Refer to the dedicated Exponentiation Operator page for syntax and code examples.

---

## 41. BigInt Support

Arbitrary-precision integers using the `n` suffix or `BigInt()` constructor. Supports arithmetic and comparison operators. Mixing BigInt and Number in the same operation throws a `TypeError`.

Refer to the dedicated BigInt Support page for syntax and code examples.

---

## 42. Promises

Full ES6 Promise API with `then`, `catch`, and `finally`. Uses a microtask queue processed automatically. Static methods include `Promise.resolve`, `Promise.reject`, `Promise.all`, and `Promise.race`.

Refer to the dedicated Promises page for syntax and code examples.

---

## 43. Generators (function*)

Functions that can be paused and resumed using `yield`. Calling a generator function returns an iterator object. Generator context is preserved across re-entrances. `yield*` delegation is supported.

Refer to the dedicated Generators page for syntax and code examples.

---

## 44. Async/Await

`async` functions return a Promise. `await` pauses execution until the Promise settles. In the AxonASP environment, `await` blocks synchronously while pumping the microtask queue. Standard `try...catch` works for error handling.

Refer to the dedicated Async/Await page for syntax and code examples.

---

## 45. Atomics

Atomic operations on SharedArrayBuffer-backed integer TypedArrays: `add`, `sub`, `and`, `or`, `xor`, `load`, `store`, `exchange`, `compareExchange`, and `isLockFree`. Strictly validates that the backing buffer is a SharedArrayBuffer.

Refer to the dedicated Atomics page for syntax and code examples.

---

## Additional Resources

Each feature documented above has its own dedicated page with complete syntax definitions, detailed remarks, and runnable code examples. Navigate to the specific subpage from the documentation menu to view the full reference.
