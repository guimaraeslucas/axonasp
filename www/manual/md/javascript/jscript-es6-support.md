# Use ES6 Features and beyond in Javascript Scripts

## Overview

AxonASP's Javascript engine supports a wide range of modern ECMAScript features, including ES6 (ES2015) additions and subsequent standards up to ES2024. This page documents all supported modern capabilities: template literals, block-scoped declarations (`let` and `const`) with Temporal Dead Zone (TDZ), arrow functions, default parameter values, rest parameters, spread in array literals, object literal shorthand, computed property names, `for...of` loops, `Object` static utilities (including `values`, `entries`, and `fromEntries`), property reflection helpers, modern `String` methods (like `includes`, `padStart`, and `at`), full Unicode support in `RegExp`, `Number` static methods, `Math` extensions, `Symbol` primitive, `Set` and `Map` collections, and a comprehensive set of `Array` utilities (including `find`, `flat`, `flatMap`, and immutable `toSorted`/`toReversed`/`toSpliced` methods).

All ES6 features described here are available in `<script runat="server" language="JScript">` blocks and in `<% language="JScript" %>` inline blocks.

---

## ECMAScript Modules (import and export)

### Syntax

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

### Remarks

- `import` and `export` are supported for server-side JavaScript modules loaded from `.js` files.
- Module loading is **synchronous**. The VM resolves and executes imported modules in the same request execution flow.
- Module instances are stored per request in a request-local module registry. The same module path executes only once per request and subsequent imports reuse the same module environment.
- Compiled module bytecode uses the global script cache. This avoids recompilation when the source did not change.
- Circular dependencies are supported with partial initialization semantics.
- Standard ASP objects (`Response`, `Request`, `Session`, `Application`, `Server`) are automatically available inside modules.
- **ReferenceError:** The VM throws a `ReferenceError` if a requested named export is missing from the source module.

---

## Weak Collections (WeakMap and WeakSet)

### Syntax

```javascript
var wm = new WeakMap();
var ws = new WeakSet();

var key = {};
wm.set(key, "data");
ws.add(key);

Response.Write(wm.get(key)); // data
Response.Write(ws.has(key)); // True
```

### Remarks

- `WeakMap` and `WeakSet` provide collections where keys (or values in `WeakSet`) are held weakly.
- **Memory Safety:** Unlike standard `Map` and `Set`, weak collections do not prevent their keys from being garbage collected. This is critical for preventing memory leaks in long-running scripts where objects are used as temporary keys.
- **Inverted Storage:** AxonASP uses an efficient "inverted storage" pattern where weak data is stored internally within the key object itself, ensuring that when the key is destroyed, the associated data is automatically reclaimed without GC overhead.
- **Valid Keys:** Objects (`{}`), functions (`function`), and **unique Symbols** (those created via `Symbol()` that are not registered in the global registry via `Symbol.for()` and are not well-known symbols like `Symbol.iterator`) can be used as keys. Attempting to use a primitive (string, number, boolean) or a restricted symbol as a key will throw a `TypeError`.
- **Non-Iterable:** Weak collections are not iterable. They do not support `for...of` loops, and they do not have `.size`, `.keys()`, `.values()`, or `.entries()` methods.

---

## Weak References (WeakRef and FinalizationRegistry)

### Syntax

```javascript
var obj = { data: 42 };

// 1. WeakRef
var wr = new WeakRef(obj);
var target = wr.deref();
if (target !== undefined) {
    Response.Write(target.data);
}

// 2. FinalizationRegistry
var registry = new FinalizationRegistry(function(heldValue) {
    // Callback executed when registered objects are garbage collected
});
registry.register(obj, "some context", obj); // register object
registry.unregister(obj); // unregister
```

### Remarks

- **WeakRef:** Provides a way to hold a weak reference to an object, function, or symbol, allowing it to be garbage collected while still attempting to access it if it hasn't been collected yet via the `deref()` method.
- **FinalizationRegistry:** Allows you to register a callback to be invoked when an object is garbage collected.
- **VM Implementation Note:** AxonASP's JScript engine focuses on short-lived, high-performance HTTP request processing and does not implement a background garbage collector. Objects typically live until the end of the script execution (or request completion). Therefore, `FinalizationRegistry` callbacks will not be triggered during standard execution, but the API and validation semantics are fully implemented for compatibility with modern JavaScript libraries that expect these features to exist.
- Target objects for `WeakRef` and `FinalizationRegistry.register` must be Objects (`{}`), Functions, or unique Symbols (not registered via `Symbol.for()`). Passing primitives will result in a `TypeError`.

---

## Block-Scoped Declarations (let and const)

### Syntax

```javascript
let x = 10;
const y = 20;

{
    let x = 30; // Shadows outer x
    const y = 40; // Shadows outer y
}
```

### Remarks

- `let` and `const` provide block-level scoping. Variables declared inside a `{}` block are only accessible within that block.
- **Temporal Dead Zone (TDZ):** Unlike `var`, accessing a `let` or `const` variable before its declaration line in the execution flow results in a `ReferenceError`.
- `const` bindings are immutable; attempting to reassign a value to a `const` variable results in a `TypeError`.

### Code Example

```javascript
<%
let a = 1;
{
    // Response.Write(a); // This would throw ReferenceError due to TDZ if 'let a' exists below
    let a = 2;
    Response.Write(a); // Output: 2
}
Response.Write(a); // Output: 1

const PI = 3.14;
// PI = 3.15; // This would throw TypeError
%>
```

---

## Explicit Resource Management (using)

### Syntax

```javascript
using resource = expression;
async using asyncResource = expression;
```

### Remarks

- `using` is supported as a block-scoped declaration and requires an initializer.
- Resource cleanup uses symbol-based disposal methods:
  - `Symbol.dispose` for `using`
  - `Symbol.asyncDispose` for `async using`
- Multiple `using` declarations in the same scope are disposed in reverse declaration order.
- Disposal runs at normal scope exit and during exception unwinding (`throw`) for the same scope.
- `async using` currently invokes `Symbol.asyncDispose` synchronously (without awaiting Promise settlement).

### Code Example

```javascript
<script runat="server" language="JScript">
var trace = [];

var firstResource = {
    [Symbol.dispose]: function() {
        trace.push("dispose:first");
    }
};

var secondResource = {
    [Symbol.dispose]: function() {
        trace.push("dispose:second");
    }
};

{
    using first = firstResource;
    using second = secondResource;
    trace.push("inside");
}

Response.Write(trace.join("|"));
// Output: inside|dispose:second|dispose:first
</script>
```

---

## Full Unicode Support

### String Code Point Escapes

ES6 introduces a new escape sequence for Unicode characters that allows representing any character using its code point value in hexadecimal between braces.

#### Syntax

```javascript
var s = "\u{1D306}"; // Tetragram for Centre
```

#### Remarks

- Supports values from `0` up to `0x10FFFF`.
- Correctly handles surrogate pairs internally. A character like `\u{1D306}` has a `.length` of 2 in JScript (representing two UTF-16 code units).

### RegExp /u flag

The `u` flag (Unicode) enables advanced Unicode features in regular expressions.

#### Syntax

```javascript
var re = /^\u{1D306}$/u;
```

#### Remarks

- When the `u` flag is present, `.` matches a full Unicode code point (even if it spans multiple UTF-16 code units).
- Enables `\u{...}` escape sequences inside the regular expression pattern.

### Code Example

```javascript
<%
// String length with surrogate pairs
var s = "\u{1D306}";
Response.Write(s.length); // Output: 2

// Unicode RegExp matching
var re = /^.$/u;
Response.Write(re.test(s)); // Output: true (matches the whole surrogate pair)
%>
```

---

## Modern Regular Expressions (RegExp)

AxonASP uses a PCRE-compatible engine for JScript Regular Expressions, supporting advanced features introduced in ES6 and subsequent standards (ES2018+).

### Named Capture Groups

Named capture groups allow you to assign names to capturing groups, which can then be accessed via the `groups` property of the match result.

#### Syntax

```javascript
var re = /(?<name>pattern)/;
```

### Lookaround Assertions

Lookaround assertions (lookahead and lookbehind) allow matching a pattern based on what precedes or follows it, without including those characters in the match.

#### Syntax

- **Positive Lookahead:** `(?=...)`
- **Negative Lookahead:** `(?!...)`
- **Positive Lookbehind:** `(?<=...)`
- **Negative Lookbehind:** `(?<!...)`

### Sticky Flag (y)

The `y` flag indicates that the match must start exactly at the `lastIndex` property of the regular expression object. If the match fails, `lastIndex` is reset to `0`.

### RegExp.prototype.flags

The `flags` property returns a string containing the flags of the regular expression object, sorted alphabetically (`g`, `i`, `m`, `s`, `u`, `y`).

### Code Example

```javascript
<script runat="server" language="JScript">
// 1. Named Capture Groups
var re = /(?<year>\d{4})-(?<month>\d{2})-(?<day>\d{2})/;
var match = re.exec("2026-05-14");
Response.Write(match.groups.year); // Output: 2026

// 2. Lookbehind
var reLookbehind = /(?<=\$)\d+/;
Response.Write(reLookbehind.exec("Price: $100")[0]); // Output: 100

// 3. Sticky Flag
var reSticky = /a/y;
reSticky.lastIndex = 1;
Response.Write(reSticky.exec("ba") !== null); // Output: true

// 4. Flags Property
Response.Write(/abc/gimuy.flags); // Output: gimuy
</script>
```

---

## Template Literals

### Syntax

```javascript
var result = `static text ${expression} more static text`;
```

Template literals are enclosed in backticks (`` ` ``). They support embedded expressions using `${expression}` placeholders and preserve literal newlines.

### Remarks

- All `${expression}` placeholders are evaluated at runtime and coerced to strings using standard JScript string coercion.
- Multiple expressions can be embedded in a single template literal.
- Multi-line template literals preserve embedded newline characters.
- Tagged template literals are not supported. A tagged template (e.g., `` tag`...` ``) resolves to `undefined`.

### Code Example

```javascript
<%
var name = "World";
var count = 42;
var msg = `Hello, ${name}! You have ${count} messages.`;
Response.Write(msg);
// Output: Hello, World! You have 42 messages.

var a = 3, b = 4;
Response.Write(`Sum: ${a + b}`);
// Output: Sum: 7
%>
```

---

## Arrow Functions

### Syntax

```javascript
// Concise body (expression result is implicitly returned)
var fn = (param1, param2) => expression;

// Block body
var fn = (param1, param2) => {
    // statements
    return value;
};
```

### Remarks

- Arrow functions do not create their own `this` binding. The value of `this` is captured **lexically** from the enclosing scope at the time the arrow function is created. This is useful for callbacks inside constructor methods.
- Arrow functions cannot be used as constructors. Using `new` with an arrow function is not supported.
- Single-parameter arrow functions without parentheses (e.g., `x => x * 2`) are supported.
- Arrow functions have an `arguments` object bound to the enclosing function's `arguments`, not their own.

### Code Example

```javascript
<script runat="server" language="JScript">
// Concise arrow function
var square = (x) => x * x;
Response.Write(square(5));
// Output: 25

// Lexical this in a constructor
function Timer() {
    this.seconds = 0;
    this.tick = function() {
        var increment = () => { this.seconds = this.seconds + 1; };
        increment();
    };
}
var t = new Timer();
t.tick();
t.tick();
Response.Write(t.seconds);
// Output: 2
</script>
```

---

## Default Parameter Values

### Syntax

```javascript
function greet(name, message = "Hello") {
    return message + ", " + name + "!";
}
```

### Remarks

- Native default parameter syntax is supported (for example, `function f(a = 10)`).
- The classic guard pattern `if (x === undefined) x = ...` is still supported and remains useful for compatibility-oriented scripts.

### Code Example

```javascript
<script runat="server" language="JScript">
function multiply(a, b = 2) {
    return a * b;
}
Response.Write(multiply(5));      // Output: 10
Response.Write(multiply(5, 3));   // Output: 15
</script>
```

---

## Tail Call Optimization (TCO)

### Syntax

```javascript
function sum(n, acc) {
    if (n === 0) {
        return acc;
    }
    return sum(n - 1, acc + n);
}
```

### Remarks

- Tail-position calls in `return` statements are optimized by the JScript VM to reuse the active function frame.
- The optimization currently applies to direct calls (`return fn(...)`) and member calls (`return obj.fn(...)`).
- Tail-call optimization is intentionally disabled when the `return` statement is inside `try`, `catch`, or `finally` blocks to preserve exception-handler semantics.
- If the tail-position call target resolves to a native host function, the VM executes it as a normal call and returns the result without frame reuse.

### Code Example

```javascript
<script runat="server" language="JScript">
function sum(n, acc) {
    if (n === 0) {
        return acc;
    }
    return sum(n - 1, acc + 1);
}

Response.Write(sum(100000, 0));
// Output: 100000
</script>
```

---

## Rest Parameters

### Syntax

```javascript
function fn(first, second, ...rest) {
    // rest is a standard array of remaining arguments
}
```

### Remarks

- The rest parameter must be the last parameter in the function signature.
- `rest` is a standard JScript array and supports all array methods.
- Only one rest parameter is allowed per function.

### Code Example

```javascript
<script runat="server" language="JScript">
function pack(head, ...rest) {
    return head + ":" + rest.length;
}
Response.Write(pack("h", 1, 2, 3));
// Output: h:3
</script>
```

---

## Object Literal Property Shorthand

### Syntax

```javascript
var x = 10;
var y = 20;
var point = { x, y }; // equivalent to { x: x, y: y }
```

### Remarks

- Shorthand property syntax is supported when the variable name and the property name are identical.
- Method shorthand (e.g., `{ greet() {} }`) follows the same rule and is available as well.

### Code Example

```javascript
<script runat="server" language="JScript">
var x = 10;
var y = 20;
var p = { x, y };
Response.Write(p.x + "," + p.y);
// Output: 10,20
</script>
```

---

## Spread in Array Literals

### Syntax

```javascript
var out = [1, 2, ...otherArray, 5];
```

### Remarks

- Spread in array literals expands one source array-like value into individual elements.
- `null` and `undefined` spread sources raise a JScript `TypeError`.
- Evaluation order is preserved left to right.

### Code Example

```javascript
<script runat="server" language="JScript">
var src = [3, 4];
var out = [1, 2, ...src, 5];
Response.Write(out.join(","));
// Output: 1,2,3,4,5
</script>
```

---

## Object Static Utilities

The following `Object` static methods are available.

### `Object.assign(target, ...sources)`

Copies enumerable own properties from each source object into `target`, from left to right, and returns `target`.

### `Object.keys(object)`

Returns an array of enumerable own property names.

### `Object.values(object)`

Returns an array of enumerable own property values.

### `Object.entries(object)`

Returns an array where each item is a two-element `[key, value]` pair for each enumerable own property.

### `Object.fromEntries(iterable)`

Converts an iterable of key-value pairs (such as an array of `[key, value]` arrays) into a new object.

### `Object.is(value1, value2)`

Returns `true` when both values are the same according to ECMAScript `Object.is` semantics. `NaN` compares equal to `NaN`, and `+0` and `-0` compare as different values.

### `Object.setPrototypeOf(object, prototype)`

Changes the prototype of `object` to `prototype` when the object is extensible. Throws a `TypeError` if the target is not an object or if its prototype cannot be changed.

### Remarks

- `Object.assign` skips `null` and `undefined` sources.
- `Object.keys`, `Object.values`, and `Object.entries` throw a JScript `TypeError` when called with `null` or `undefined`.
- Return values are standard JScript arrays and are compatible with existing array operations.
- Symbol-keyed properties are intentionally excluded from `Object.keys`, `Object.values`, and `Object.entries` to reduce collision risks in legacy code.

### `Object.getOwnPropertySymbols(object)`

Returns an array of the object's own symbol-keyed properties in symbol form.

- `Object.getOwnPropertySymbols` ignores string-keyed properties.
- Prototype inheritance is not included; only own properties are reported.

### Code Example

```javascript
<script runat="server" language="JScript">
var target = { a: 1 };
Object.assign(target, { b: 2 }, { c: 3 });

Response.Write(Object.keys(target).join(","));
// Output: a,b,c

Response.Write(Object.values(target).join(","));
// Output: 1,2,3

var e = Object.entries(target);
Response.Write(e[0][0] + ":" + e[0][1]);
// Output: a:1

var entries = [["x", 10], ["y", 20]];
var obj = Object.fromEntries(entries);
Response.Write(obj.x + "," + obj.y);
// Output: 10,20

var s = Symbol("id");
var o = {};
Object.setPrototypeOf(o, { inherited: 1 });
o[s] = 99;
Response.Write(Object.is(NaN, NaN));
// Output: true
Response.Write(Object.getOwnPropertySymbols(o).length);
// Output: 1
</script>
```

---

## Property Reflection Helpers

### `Object.getOwnPropertyDescriptor(object, propertyName)`

Returns the property descriptor for an own property of `object`. The descriptor object contains the following fields: `value`, `writable`, `enumerable`, and `configurable`.

### `Object.getOwnPropertyDescriptors(object)`

Returns an object whose own properties are the property descriptors for all own properties of `object`. Each key maps to the same descriptor structure returned by `Object.getOwnPropertyDescriptor`.

### Remarks

- Both methods operate only on own properties. Inherited properties are not reported.
- Symbol-keyed internals follow the same visibility constraints as `Object.keys` and are not included in the result.
- `Object.defineProperty` is available and can be used to define non-enumerable or read-only properties before inspecting them with these helpers.

### Code Example

```javascript
<script runat="server" language="JScript">
var o = {};
Object.defineProperty(o, "hidden", {
    value: 10,
    writable: false,
    enumerable: false,
    configurable: false
});
var d = Object.getOwnPropertyDescriptor(o, "hidden");
var all = Object.getOwnPropertyDescriptors(o);
Response.Write(d.value + "|" + all.hidden.writable);
// Output: 10|false
</script>
```

---

## Array Search Utilities

### `Array.prototype.find(callback[, thisArg])`

Returns the first element that satisfies `callback`. Returns `undefined` when no element matches.

### `Array.prototype.findIndex(callback[, thisArg])`

Returns the index of the first element that satisfies `callback`. Returns `-1` when no element matches.

### Code Example

```javascript
<script runat="server" language="JScript">
var arr = [3, 7, 11, 14];
Response.Write(arr.find(function (x) { return x > 10; }));
// Output: 11
Response.Write(arr.findIndex(function (x) { return x > 10; }));
// Output: 2
</script>
```

---

## Array Construction Utilities

### `Array.from(arrayLike[, mapFn])`

Converts an array-like or iterable object into a standard JScript array. Accepts an optional mapping function that is applied to each element.

### `Array.of(...items)`

Creates a new array from its arguments. Unlike `new Array(n)`, `Array.of(n)` always creates a one-element array containing `n`.

### Code Example

```javascript
<script runat="server" language="JScript">
var a = Array.from({ length: 2, 0: "x", 1: "y" });
var b = Array.of(7, 8, 9);
Response.Write(a.join("-") + "|" + b.join("-"));
// Output: x-y|7-8-9
</script>
```

---

## Array In-place Operations

### `Array.prototype.fill(value[, start[, end]])`

Fills all elements from `start` to `end` (exclusive) with `value`, in place. Negative indices are resolved relative to the array length. Returns the modified array.

### `Array.prototype.copyWithin(target[, start[, end]])`

Copies a portion of the array (from `start` to `end`, exclusive) to another position (`target`) within the same array, in place. Does not change the array length. Returns the modified array.

### `Array.prototype.keys()`

Returns an Array Iterator object containing each numeric index key from the array.

### `Array.prototype.entries()`

Returns an Array Iterator object containing `[index, value]` pairs for each array element.

### `Array.prototype.at(index)`

Returns the element at the specified `index`. Supports relative indexing from the end if `index` is negative.

### `Array.prototype.flat([depth])`

Returns a new array with all sub-array elements concatenated into it recursively up to the specified `depth`. Defaults to `1`.

### `Array.prototype.flatMap(callback[, thisArg])`

Returns a new array formed by applying a given callback function to each element of the array, and then flattening the result by one level.

### `Array.prototype.toSorted([compareFn])`

Returns a **new** array with the elements sorted in ascending order. Unlike `sort()`, it does not mutate the original array.

### `Array.prototype.toReversed()`

Returns a **new** array with the elements in reversed order. Unlike `reverse()`, it does not mutate the original array.

### `Array.prototype.toSpliced(start[, deleteCount[, ...items]])`

Returns a **new** array with some elements removed and/or replaced at a given index. Unlike `splice()`, it does not mutate the original array.

### Remarks

- Methods like `fill` and `copyWithin` operate in place and return the same array reference.
- `keys()` and `entries()` return standard iterable Array Iterator objects and can be consumed by `for...of`.
- Modern immutable methods (`toSorted`, `toReversed`, `toSpliced`) always return a new array instance.
- Negative index arguments in `at`, `fill`, and `copyWithin` are normalized relative to the array length.

### Code Example

```javascript
<script runat="server" language="JScript">
var arr = [1, [2, 3]];
Response.Write(JSON.stringify(arr.flat()));
// Output: [1,2,3]

var original = [3, 1, 2];
var sorted = original.toSorted();
Response.Write(sorted.join(","));
// Output: 1,2,3
Response.Write(original.join(","));
// Output: 3,1,2 (unchanged)

for (var k of [10, 20].keys()) {
    Response.Write(k + " ");
}
// Output: 0 1

for (var e of [10, 20].entries()) {
    Response.Write(e[0] + ":" + e[1] + " ");
}
// Output: 0:10 1:20

Response.Write("abc".at(-1));
// Output: c
</script>
```

---

## ES6 String Methods

The following methods are available on `String` values.

### `String.prototype.includes(searchString[, position])`

Returns `true` if `searchString` is found anywhere within the string at or after `position` (default `0`); `false` otherwise. Case-sensitive. Raises a `TypeError` if `searchString` is a `RegExp`.

### `String.prototype.startsWith(searchString)`

Returns `true` if the string begins with `searchString`; `false` otherwise. Case-sensitive.

### `String.prototype.endsWith(searchString)`

Returns `true` if the string ends with `searchString`; `false` otherwise. Case-sensitive.

### `String.prototype.repeat(count)`

Returns a new string containing `count` repetitions of the original string. Returns an empty string if `count` is 0.

### `String.prototype.at(index)`

Returns the character at the specified `index`. Supports relative indexing from the end if `index` is negative.

### `String.prototype.codePointAt(position)`

Returns the Unicode code point value at `position`. If `position` is out of range, returns `undefined`.

### `String.prototype.normalize([form])`

Returns the Unicode Normalization Form of the string. Supported values are `NFC`, `NFD`, `NFKC`, and `NFKD`. If omitted, `NFC` is used.

### `String.prototype.padStart(targetLength, padString)`

Pads the string from the start with `padString` until the total length reaches `targetLength`. If `padString` is not supplied, spaces are used.

### `String.prototype.padEnd(targetLength, padString)`

Pads the string from the end with `padString` until the total length reaches `targetLength`. If `padString` is not supplied, spaces are used.

### Code Example

```javascript
<script runat="server" language="JScript">
var s = "Hello World";

Response.Write(s.includes("World"));        // Output: true
Response.Write(s.includes("World", 6));     // Output: true
Response.Write(s.startsWith("Hello"));      // Output: true
Response.Write(s.endsWith("World"));        // Output: true
Response.Write("ab".repeat(3));             // Output: ababab
Response.Write("5".padStart(3, "0"));       // Output: 005
Response.Write("5".padEnd(3, "0"));         // Output: 500
Response.Write("A😀B".codePointAt(1));       // Output: 128512
Response.Write("e\u0301".normalize("NFC")); // Output: é

var regexError = false;
try {
    "hello".includes(new RegExp("h"));
} catch (e) {
    regexError = String(e).indexOf("TypeError") !== -1;
}
Response.Write(regexError);                 // Output: true
</script>
```

---

## ES6 Number Static Methods

The following static methods are available on the `Number` object.

### `Number.isInteger(value)`

Returns `true` only if `value` is a number with no fractional part and is not `Infinity` or `NaN`. Does **not** coerce non-number values; non-numbers return `false`.

### `Number.isNaN(value)`

Returns `true` only if `value` is the numeric `NaN`. Does **not** coerce non-number values; non-numbers always return `false`. This differs from the global `isNaN()` function, which coerces its argument.

### `Number.isFinite(value)`

Returns `true` only if `value` is a finite number. Does **not** coerce non-number values; non-numbers always return `false`.

### `Number.isSafeInteger(value)`

Returns `true` if `value` is an integer in the range `-(2^53 - 1)` to `2^53 - 1` inclusive, and has no fractional part. Does **not** coerce non-number values.

### `Number.parseInt(string, radix)`

Equivalent to the global `parseInt()` function. Parses `string` as an integer in the specified `radix` (2–36). Defaults to base 10.

### `Number.parseFloat(string)`

Equivalent to the global `parseFloat()` function. Parses `string` as a floating-point number.

### Number Constants

The `Number` object exposes the following read-only constants:

| Constant | Value |
|---|---|
| `Number.MAX_SAFE_INTEGER` | 9007199254740991 |
| `Number.MIN_SAFE_INTEGER` | -9007199254740991 |
| `Number.MAX_VALUE` | ~1.7976931348623157e+308 |
| `Number.MIN_VALUE` | ~5e-324 |
| `Number.EPSILON` | ~2.220446049250313e-16 |
| `Number.POSITIVE_INFINITY` | `Infinity` |
| `Number.NEGATIVE_INFINITY` | `-Infinity` |
| `Number.NaN` | `NaN` |

### Code Example

```javascript
<script runat="server" language="JScript">
Response.Write(Number.isInteger(42));          // Output: true
Response.Write(Number.isInteger(42.5));        // Output: false
Response.Write(Number.isInteger("42"));        // Output: false

Response.Write(Number.isNaN(NaN));             // Output: true
Response.Write(Number.isNaN(42));              // Output: false
Response.Write(Number.isNaN("NaN"));           // Output: false

Response.Write(Number.isFinite(100));          // Output: true
Response.Write(Number.isFinite(Infinity));     // Output: false

Response.Write(Number.isSafeInteger(9007199254740991));  // Output: true
Response.Write(Number.isSafeInteger(9007199254740992));  // Output: false

Response.Write(Number.MAX_SAFE_INTEGER);       // Output: 9007199254740991
Response.Write(Number.EPSILON);                // Output: 2.220446049250313e-16
</script>
```

---

## Binary and Octal Numeric Literals

### Syntax

```javascript
var b = 0b1010; // binary
var o = 0o744;  // octal
```

### Remarks

- Prefix `0b` or `0B` parses base-2 integer literals.
- Prefix `0o` or `0O` parses base-8 integer literals.

### Code Example

```javascript
<script runat="server" language="JScript">
Response.Write(0b1010); // Output: 10
Response.Write(0o744);  // Output: 484
</script>
```

---

## Global URI Functions

The following URI helper functions are available globally.

### `encodeURI(uri)`

Encodes a complete URI string while preserving URI-reserved separators such as `:`, `/`, `?`, `&`, `=`, and `#`.

### `decodeURI(uri)`

Decodes a complete URI string. Reserved separators remain preserved when they were percent-encoded.

### `encodeURIComponent(component)`

Encodes a URI component (such as one query value) and escapes reserved characters like `=`, `&`, and `+`.

### `decodeURIComponent(component)`

Decodes an encoded URI component.

### Code Example

```javascript
<script runat="server" language="JScript">
var full = "https://example.com/a path/?q=hello world&x=1+2#frag";
Response.Write(encodeURI(full));
// Output: https://example.com/a%20path/?q=hello%20world&x=1+2#frag

var component = "q=hello world&x=1+2";
var encoded = encodeURIComponent(component);
Response.Write(encoded);
// Output: q%3Dhello%20world%26x%3D1%2B2

Response.Write(decodeURIComponent(encoded));
// Output: q=hello world&x=1+2
</script>
```

---

## Math Extensions

The following additional methods are available on the `Math` object.

### `Math.trunc(x)`

Returns the integer part of `x` by removing the fractional digits.

### `Math.sign(x)`

Returns `1` for positive values, `-1` for negative values, and `0` for zero. Returns `NaN` for `NaN` input.

### `Math.cbrt(x)`

Returns the cube root of `x`.

### Additional Methods

- `Math.acosh(x)`
- `Math.asinh(x)`
- `Math.atanh(x)`
- `Math.expm1(x)`
- `Math.log1p(x)`
- `Math.log10(x)`
- `Math.log2(x)`
- `Math.hypot(...values)`
- `Math.fround(x)`
- `Math.imul(a, b)`
- `Math.clz32(x)`

### Code Example

```javascript
<script runat="server" language="JScript">
Response.Write(Math.trunc(4.9)); // Output: 4
Response.Write(Math.sign(-12));  // Output: -1
Response.Write(Math.cbrt(27));   // Output: 3
Response.Write(Math.hypot(3, 4)); // Output: 5
Response.Write(Math.imul(0xffffffff, 5)); // Output: -5
Response.Write(Math.clz32(1)); // Output: 31
</script>
```

---

## Symbol Primitive

### Syntax

```javascript
var sym = Symbol(description);
```

### Remarks

- Each call to `Symbol()` returns a unique value that is never equal to any other `Symbol` or primitive.
- Symbols can be used as object property keys to create collision-safe identifiers.
- Calling `new Symbol()` raises a `TypeError`. `Symbol` is not a constructor.
- Symbol-keyed properties are intentionally hidden from `Object.keys`, `Object.values`, and `Object.entries` to prevent unintended exposure in enumeration.

### Code Example

```javascript
<script runat="server" language="JScript">
var s1 = Symbol("id");
var s2 = Symbol("id");
var o = {};
o[s1] = 42;
Response.Write((s1 !== s2) + "|" + o[s1]);
// Output: true|42
</script>
```

---

## Symbol Primitive — Well-Known Symbols and Global Registry

Well-known symbols are pre-defined `Symbol` values stored as properties of the `Symbol` constructor object.

| Symbol | Description |
|---|---|
| `Symbol.iterator` | Default iterator for `for...of` loops |
| `Symbol.toStringTag` | Object `[object X]` tag override |
| `Symbol.species` | Species constructor for derived objects |
| `Symbol.hasInstance` | Custom `instanceof` behavior |
| `Symbol.toPrimitive` | Custom primitive conversion |

The global symbol registry allows sharing symbols across realms via `Symbol.for` and `Symbol.keyFor`.

### Code Example

```javascript
<script runat="server" language="JScript">
// Well-known symbols are of type "symbol"
Response.Write(typeof Symbol.iterator);   // Output: symbol
Response.Write(typeof Symbol.toStringTag); // Output: symbol

// Symbol.for — global registry: same key returns same symbol
var a = Symbol.for("appToken");
var b = Symbol.for("appToken");
Response.Write(a === b); // Output: true

// Symbol.keyFor — retrieve key from registry
Response.Write(Symbol.keyFor(a)); // Output: appToken

// Locally created symbols are NOT in the registry
var local = Symbol("local");
Response.Write(Symbol.keyFor(local) === undefined); // Output: true
</script>
```

---

## Iteration Protocol — `for...of` and Custom Iterables

The iteration protocol allows JScript objects to define or customize their iteration behavior, such as which values are looped over in a `for...of` construct.

### `for...of` Statement

The `for...of` statement creates a loop iterating over iterable objects, including built-in `Array`, `String`, `Set`, `Map`, and custom iterables.

### Built-in Iterables

- **Array**: Iterates over elements.
- **String**: Iterates over characters (handling surrogate pairs).
- **Set**: Iterates over unique values.
- **Map**: Iterates over `[key, value]` entries.

### Custom Iterables

To make an object iterable, it must implement the `[Symbol.iterator]` method, which returns an **Iterator** object. An iterator is an object that has a `next()` method returning an object with two properties: `value` (the next value) and `done` (a boolean indicating completion).

### Code Example

```javascript
<script runat="server" language="JScript">
// 1. Iterate over an Array
var fruits = ["Apple", "Orange", "Banana"];
for (var fruit of fruits) {
    Response.Write(fruit + " "); // Output: Apple Orange Banana 
}

// 2. Manual Iterator usage
var it = fruits[Symbol.iterator]();
var res = it.next();
while (!res.done) {
    Response.Write(res.value + " ");
    res = it.next();
}

// 3. Custom Iterable
var range = {
    from: 1,
    to: 3,
    [Symbol.iterator]: function() {
        return {
            current: this.from,
            last: this.to,
            next: function() {
                if (this.current <= this.last) {
                    return { value: this.current++, done: false };
                } else {
                    return { value: undefined, done: true };
                }
            }
        };
    }
};

for (var n of range) {
    Response.Write(n + " "); // Output: 1 2 3
}
</script>
```

---

## Binary Data — ArrayBuffer and Typed Arrays

### Syntax

```javascript
var buffer = new ArrayBuffer(byteLength);
var view   = new Uint8Array(buffer);
var view   = new Uint8Array(length);
var view   = new Uint8Array([1, 2, 3]);
var dv     = new DataView(buffer [, byteOffset [, byteLength]]);
```

### Remarks

- `ArrayBuffer` holds a raw byte block. Its `byteLength` property returns its size in bytes. Use `ArrayBuffer.isView(v)` to test whether a value is a typed array view.
- **Typed arrays** provide strongly-typed views over an `ArrayBuffer`. All supported types are listed in the table below.
- `DataView` gives byte-level control over reads and writes including explicit endianness.
- Typed array constructors can be called with: a byte **length**, an existing **ArrayBuffer**, or an **array-like** source (plain array or another typed array).
- Index reads past the end of the view return `undefined`. Index writes past the end are silently ignored.
- Calling a typed array constructor without `new` raises a `TypeError`.

### Supported Typed Array Types

| Constructor | Element type | Bytes per element |
|---|---|---|
| `Int8Array` | Signed 8-bit integer | 1 |
| `Uint8Array` | Unsigned 8-bit integer | 1 |
| `Uint8ClampedArray` | Unsigned 8-bit integer, clamped [0–255] | 1 |
| `Int16Array` | Signed 16-bit integer | 2 |
| `Uint16Array` | Unsigned 16-bit integer | 2 |
| `Int32Array` | Signed 32-bit integer | 4 |
| `Uint32Array` | Unsigned 32-bit integer | 4 |
| `Float32Array` | 32-bit IEEE 754 float | 4 |
| `Float64Array` | 64-bit IEEE 754 float | 8 |
| `BigInt64Array` | Signed 64-bit integer (BigInt) | 8 |
| `BigUint64Array` | Unsigned 64-bit integer (BigInt) | 8 |

### Typed Array Properties

| Property | Description |
|---|---|
| `length` | Number of elements |
| `byteLength` | Total size in bytes |
| `byteOffset` | Offset into the backing `ArrayBuffer` |
| `buffer` | The underlying `ArrayBuffer` |

### Typed Array Methods

| Method | Description |
|---|---|
| `set(array [, offset])` | Copy values from an array-like source |
| `subarray([begin [, end]])` | Return a new view over the same buffer |
| `fill(value [, start [, end]])` | Fill all or part of the view with a value |
| `slice([begin [, end]])` | Return a new typed array copy of the range |
| `forEach(callback)` | Iterate over each element |
| `indexOf(value [, fromIndex])` | Return first index of a matching value, or -1 |

### ArrayBuffer Methods

| Method | Description |
|---|---|
| `slice([begin [, end]])` | Return a new `ArrayBuffer` containing a copy of the byte range |
| `ArrayBuffer.isView(value)` | Return `true` if the value is a typed array or DataView |

### DataView Methods

| Method | Description |
|---|---|
| `getInt8(offset)` | Read signed 8-bit int |
| `getUint8(offset)` | Read unsigned 8-bit int |
| `getInt16(offset [, littleEndian])` | Read signed 16-bit int |
| `getUint16(offset [, littleEndian])` | Read unsigned 16-bit int |
| `getInt32(offset [, littleEndian])` | Read signed 32-bit int |
| `getUint32(offset [, littleEndian])` | Read unsigned 32-bit int |
| `getFloat32(offset [, littleEndian])` | Read 32-bit float |
| `getFloat64(offset [, littleEndian])` | Read 64-bit float |
| `setInt8(offset, value)` | Write signed 8-bit int |
| `setUint8(offset, value)` | Write unsigned 8-bit int |
| `setInt16(offset, value [, littleEndian])` | Write signed 16-bit int |
| `setUint16(offset, value [, littleEndian])` | Write unsigned 16-bit int |
| `setInt32(offset, value [, littleEndian])` | Write signed 32-bit int |
| `setUint32(offset, value [, littleEndian])` | Write unsigned 32-bit int |
| `setFloat32(offset, value [, littleEndian])` | Write 32-bit float |
| `setFloat64(offset, value [, littleEndian])` | Write 64-bit float |

### Code Example

```javascript
<script runat="server" language="JScript">
// --- ArrayBuffer and Uint8Array ---
var buffer = new ArrayBuffer(4);
var view = new Uint8Array(buffer);
view[0] = 10;
view[1] = 20;
view[2] = 30;
view[3] = 40;
Response.Write(view[0] + "," + view[1] + "," + view[2] + "," + view[3]);
// Output: 10,20,30,40

// --- Uint8ClampedArray ---
var clamped = new Uint8ClampedArray(2);
clamped[0] = 300; // clamped to 255
clamped[1] = -5;  // clamped to 0
Response.Write(clamped[0] + "," + clamped[1]);
// Output: 255,0

// --- Int32Array from plain array ---
var ints = new Int32Array([-100, 0, 100]);
Response.Write(ints[0] + "," + ints.byteLength);
// Output: -100,12

// --- DataView with explicit endianness ---
var db = new ArrayBuffer(8);
var dv = new DataView(db);
dv.setInt32(0, 0xDEADBEEF, false); // big-endian
Response.Write(dv.getInt32(0, false));
// Output: -559038737

// --- ArrayBuffer.slice ---
var sliced = buffer.slice(1, 3);
Response.Write(sliced.byteLength);
// Output: 2

// --- for...of on typed array ---
var a = new Uint8Array([10, 20, 30]);
var sum = 0;
for (var v of a) { sum += v; }
Response.Write(sum);
// Output: 60
</script>
```

---

## Set and Map Collections

### `Set`

A `Set` stores unique values. Duplicate values are silently ignored on insertion.

| Method | Description |
|---|---|
| `set.add(value)` | Inserts `value` and returns the `Set`. |
| `set.has(value)` | Returns `true` if `value` is present. |
| `set.delete(value)` | Removes `value`. Returns `true` if the value existed. |
| `set.clear()` | Removes all elements. |
| `set.size` | Returns the number of unique elements. |

### `Map`

A `Map` stores key/value pairs and preserves insertion order.

| Method | Description |
|---|---|
| `map.set(key, value)` | Sets the entry for `key` and returns the `Map`. |
| `map.get(key)` | Returns the value associated with `key`, or `undefined`. |
| `map.has(key)` | Returns `true` if an entry for `key` exists. |
| `map.delete(key)` | Removes the entry for `key`. Returns `true` if it existed. |
| `map.clear()` | Removes all entries. |
| `map.size` | Returns the number of entries. |

### Code Example

```javascript
<script runat="server" language="JScript">
var s = new Set();
s.add("a");
s.add("b");
s.add("a"); // duplicate, ignored
Response.Write(s.has("a") + "|" + s.size);
// Output: true|2

var m = new Map();
m.set("k", 10);
Response.Write(m.has("k") + "|" + m.get("k"));
// Output: true|10
</script>
```

---

## Computed Property Names

### Syntax

```javascript
var key = "name";
var obj = { [key]: "Alice" };
var obj2 = { [prefix + "_en"]: "Hello", ["dynamic"]: 42 };
```

Use square brackets around a key expression inside an object literal to compute the property name at runtime.

### Remarks

- The expression inside `[...]` is evaluated at runtime and coerced to a string to form the property name.
- Any valid JScript expression can be used as the key: variables, string concatenations, function calls, and so on.
- Computed keys can be mixed freely with static keys and shorthand properties in the same literal.
- Numeric computed keys are coerced to strings before assignment (consistent with JScript's property model).

### Code Example

```javascript
<script runat="server" language="JScript">
var type = "color";
var o = {
    static: "fixed",
    [type]: "red",
    [type + "_code"]: "#FF0000"
};
Response.Write(o.static);       // Output: fixed
Response.Write(o.color);        // Output: red
Response.Write(o.color_code);   // Output: #FF0000

// Dynamic method name
var methodKey = "greet";
var api = { [methodKey]: function(n) { return "Hello, " + n; } };
Response.Write(api.greet("World")); // Output: Hello, World
</script>
```

---

## Internationalization API (Intl)

### Syntax

```javascript
var dtf = new Intl.DateTimeFormat(locales[, options]);
var nfmt = new Intl.NumberFormat(locales[, options]);
```

### Remarks

- `Intl` is available as a global namespace in JScript.
- `DateTimeFormat` and `NumberFormat` use AxonASP locale profiles and the current server locale when no locale is supplied.
- Locale input can be a string or an array-like value. AxonASP uses the first usable locale tag and falls back to the effective server locale, then `en-US`.
- `Intl.DateTimeFormat` supports `dateStyle`, `timeStyle`, `year`, `month`, `day`, `weekday`, `hour`, `minute`, `second`, and `hour12`.
- `Intl.NumberFormat` supports `style: "decimal"`, `style: "currency"`, and `style: "percent"`.
- Unsupported locale values and extra options are ignored or fall back to the closest supported locale profile.

### Code Example

```javascript
<script runat="server" language="JScript">
var dateValue = new Date(Date.UTC(2026, 0, 2, 3, 4, 5));
var enDate = new Intl.DateTimeFormat("en-US", { dateStyle: "short" }).format(dateValue);
var ptDate = new Intl.DateTimeFormat("pt-BR", { dateStyle: "short" }).format(dateValue);
var deNumber = new Intl.NumberFormat("de-DE", { style: "currency", currency: "EUR", maximumFractionDigits: 2 }).format(1234567.89);

Response.Write(enDate + "\n");
Response.Write(ptDate + "\n");
Response.Write(deNumber + "\n");
// Output:
// 1/2/2026
// 02/01/2026
// € 1.234.567,89
</script>
```

---

## Destructuring Assignment

Destructuring assignment is a syntax that makes it possible to unpack values from arrays, or properties from objects, into distinct variables.

### Object Destructuring

Object destructuring allows you to extract multiple properties from an object and assign them to variables in a single statement.

#### Syntax

```javascript
var { p1, p2 } = object;
var { p1: v1, p2: v2 } = object;
var { p1 = defaultValue } = object; // Default value
var { p1, ...rest } = object; // Object rest
```

#### Remarks

- If a variable name matches a property name, you can use the shorthand `{ p1, p2 }`.
- You can map a property to a different variable name using `{ property: variable }`.
- **Default Values:** You can provide a default value using `=`. The default is only used if the property is strictly `undefined`.
- **Rest Property:** The `...rest` syntax creates a new object containing all remaining enumerable own properties of the source object.
- Nested destructuring is supported: `var { a: { b } } = obj;`.
- Computed property names can be used: `var { [key]: value } = obj;`.
- **Validation:** Attempting to destructure `null` or `undefined` raises a `TypeError`.

#### Code Example

```javascript
<script runat="server" language="JScript">
var user = { id: 1, name: "Alice", details: { age: 25 } };

// Basic destructuring
var { id, name } = user;
Response.Write(id + ": " + name + "\n"); // Output: 1: Alice

// Default values and rest
var { role = "guest", ...others } = { id: 2 };
Response.Write(role + "|" + others.id + "\n"); // Output: guest|2

// Renaming and nested
var { name: userName, details: { age } } = user;
Response.Write(userName + " is " + age + "\n"); // Output: Alice is 25

// Assignment without declaration (requires parentheses)
var x, y;
({ x, y } = { x: 10, y: 20 });
Response.Write(x + y); // Output: 30
</script>
```

### Array Destructuring

Array destructuring allows you to extract elements from arrays or any iterable object (like Strings, Sets, or Maps) using an array-like syntax.

#### Syntax

```javascript
var [a, b] = iterable;
var [a, , c] = iterable; // Elision (skipping elements)
var [a = 10] = iterable; // Default value
var [a, ...rest] = iterable; // Array rest
var [a, [b, c]] = iterable; // Nested array destructuring
```

#### Remarks

- Values are extracted in order from the source iterable.
- **Elision:** You can skip elements using commas: `var [first, , last] = [1, 2, 3];`.
- **Default Values:** Assigns a default if the iterable yields `undefined` or is exhausted.
- **Rest Elements:** The `...rest` syntax collects all remaining yielded values into a new Array.
- **Iteration Protocol:** Unlike object destructuring, array destructuring works with any object that implements the ES6 Iteration Protocol. This includes Strings (yielding characters), Maps (yielding `[key, value]` pairs), and Sets.
- **Validation:** Attempting to destructure a non-iterable value (like `true` or a plain object without `[Symbol.iterator]`) raises a `TypeError`.

#### Code Example

```javascript
<script runat="server" language="JScript">
// 1. Basic array with rest
var [first, ...others] = ["Red", "Green", "Blue"];
Response.Write(first + ":" + others.length + "\n"); // Output: Red:2

// 2. Default values
var [x = 1, y = 2] = [42];
Response.Write(x + "|" + y + "\n"); // Output: 42|2

// 3. String (iterable)
var [h, e, l, l2, o] = "Hello";
Response.Write(h + e + l + l2 + o + "\n"); // Output: Hello

// 4. Nested
var [a, [b, c]] = [1, [2, 3]];
Response.Write(a + b + c + "\n"); // Output: 6

// 5. Map (yields [key, value] pairs)
var map = new Map();
map.set("id", 42);
var [[key, val]] = map;
Response.Write(key + "=" + val + "\n"); // Output: id=42
</script>
```

---

## ES6 Classes

AxonASP supports ES6 classes for object-oriented programming. Under the hood, classes are built upon JScript's existing prototype-based inheritance model but with modern syntax and strict semantics.

### Syntax

```javascript
class MyClass [extends BaseClass] {
    constructor(...args) {
        [super(...args);]
        // initialization
    }
    
    method() { ... }
    
    static staticMethod() { ... }
    
    get property() { ... }
    set property(value) { ... }
}
```

### Remarks

- **Strict Mode:** All code within a `class` body (including methods and the constructor) implicitly runs in **Strict Mode**.
- **No Hoisting:** Unlike function declarations, classes are not hoisted. You must declare a class before you can use it (Temporal Dead Zone applies).
- **Instantiation:** Classes must be instantiated with the `new` operator. Calling a class constructor as a normal function (without `new`) throws a `TypeError`.
- **Instance Methods:** Methods defined inside the class are attached to the class's `prototype`.
- **Static Methods:** Methods marked with the `static` keyword are attached directly to the class constructor function.
- **Inheritance:** When a class uses `extends`, AxonASP evaluates the superclass, validates that it is a constructor or `null`, and wires both the constructor chain and the prototype chain.
- **Null Heritage:** `extends null` is supported. In that case, the class prototype chain terminates at `null`.
- **Accessors:** `get` and `set` syntax is supported for defining property getters and setters.
- **Private Fields:** ES2022 private class fields (e.g. `#propertyName`) and private static fields (e.g. `static #staticProperty`) are fully supported. They provide true encapsulation without external memory overhead.

### Code Example

```javascript
<script runat="server" language="JScript">
class Person {
    constructor(name) {
        this._name = name;
    }

    // Instance method
    greet() {
        return "Hello, I'm " + this._name;
    }

    // Static method
    static species() {
        return "Homo Sapiens";
    }

    // Accessors
    get name() {
        return this._name.toUpperCase();
    }
    
    set name(value) {
        this._name = value;
    }
}

var p = new Person("Alice");
Response.Write(p.greet() + "<br>");       // Output: Hello, I'm Alice
Response.Write(Person.species() + "<br>"); // Output: Homo Sapiens
Response.Write(p.name + "<br>");           // Output: ALICE

p.name = "Bob";
Response.Write(p.name);                   // Output: BOB
</script>
```

### Inheritance with super()

When a class extends another class, you can use the `super()` keyword to invoke the parent class's constructor and `super.method()` to call parent class methods.

#### super() in Derived Class Constructors

The `super()` call must be made before accessing `this` in a derived class constructor. If `this` is accessed before `super()` completes, a `ReferenceError` is thrown (Temporal Dead Zone).

```javascript
<script runat="server" language="JScript">
class Animal {
    constructor(name) {
        this.name = name;
    }
    speak() {
        return this.name + " makes a sound";
    }
}

class Dog extends Animal {
    constructor(name, breed) {
        super(name);        // Call parent constructor
        this.breed = breed;
    }
    speak() {
        return super.speak() + " - woof!";
    }
}

var dog = new Dog("Buddy", "Golden Retriever");
Response.Write(dog.speak()); // Output: Buddy makes a sound - woof!
</script>
```

#### super.method() Calls

Use `super.method()` to invoke a method from the parent class. This is useful for extending parent behavior without completely overriding it.

```javascript
<script runat="server" language="JScript">
class Calculator {
    add(a, b) {
        return a + b;
    }
}

class AdvancedCalculator extends Calculator {
    add(a, b) {
        var result = super.add(a, b);
        return result + 10; // Add 10 to the base result
    }
}

var calc = new AdvancedCalculator();
Response.Write(calc.add(5, 3)); // Output: 18 (5 + 3 + 10)
</script>
```

#### super Property Access

You can also use `super` to set or access properties on the parent class prototype:

```javascript
<script runat="server" language="JScript">
class Base {
    greet() { return "Hello"; }
}

class Derived extends Base {
    greet() {
        return super.greet() + " World";
    }
    setData(val) {
        super.data = val; // Set on instance via parent
    }
}

var d = new Derived();
Response.Write(d.greet());   // Output: Hello World
d.setData(42);
Response.Write(d.data);      // Output: 42
</script>
```

#### Remarks

- `super()` **must** be called in a derived class constructor before accessing `this`. Accessing `this` before `super()` throws a `ReferenceError`.
- `super.method()` resolves the method from the parent class's prototype and calls it with the current `this` context.
- Multi-level inheritance is fully supported: `class C extends B extends A` works as expected, with each level able to call its parent via `super`.
- Static methods cannot use `super.method()` unless they are inside a derived static method that explicitly calls a parent static method.

---

## Optional Chaining (?.)

### Syntax

```javascript
obj?.property
obj?.[expression]
obj?.method()
```

### Remarks

- The optional chaining operator (`?.`) allows reading the value of a property located deep within a chain of connected objects without having to expressly validate that each reference in the chain is valid.
- If the object before the `?.` is `null` or `undefined`, the expression short-circuits and returns `undefined` instead of throwing an error.
- Works for property access, bracket access, and function calls.

### Code Example

```javascript
<script runat="server" language="JScript">
var user = { info: { name: "Alice" } };
Response.Write(user?.info?.name); // Output: Alice
Response.Write(user?.settings?.theme); // Output: undefined (no error)

var fn = null;
Response.Write(fn?.()); // Output: undefined (no error)
</script>
```

---

## Nullish Coalescing (??)

### Syntax

```javascript
var result = leftExpr ?? rightExpr;
```

### Remarks

- The nullish coalescing operator (`??`) is a logical operator that returns its right-hand side operand when its left-hand side operand is `null` or `undefined`, and otherwise returns its left-hand side operand.
- Unlike the OR operator (`||`), it does not return the right-hand side for other "falsy" values like `0`, `""`, or `false`.

### Code Example

```javascript
<script runat="server" language="JScript">
Response.Write(null ?? "default"); // Output: default
Response.Write(undefined ?? "default"); // Output: default
Response.Write(0 ?? 42); // Output: 0
Response.Write("" ?? "hello"); // Output: (empty string)
Response.Write(false ?? true); // Output: False
</script>
```

---

## Logical Assignment (||=, &&=, ??=)

### Syntax

```javascript
a ||= b;  // Logical OR assignment
a &&= b;  // Logical AND assignment
a ??= b;  // Nullish coalescing assignment
```

### Remarks

- `a ||= b` only assigns `b` to `a` if `a` is falsy.
- `a &&= b` only assigns `b` to `a` if `a` is truthy.
- `a ??= b` only assigns `b` to `a` if `a` is nullish (`null` or `undefined`).
- These operators short-circuit; the right-hand side is only evaluated if the assignment condition is met.

### Code Example

```javascript
<script runat="server" language="JScript">
var a = 0;
a ||= 10;
Response.Write(a); // Output: 10

var b = 5;
b &&= 20;
Response.Write(b); // Output: 20

var c = null;
c ??= 30;
Response.Write(c); // Output: 30
</script>
```

---

## Exponentiation Operator (**)

### Syntax

```javascript
var result = base ** exponent;
var a **= exponent;
```

### Remarks

- The exponentiation operator (`**`) returns the result of raising the first operand to the power of the second operand.
- It is equivalent to `Math.pow()`, but also supports `BigInt`.

### Code Example

```javascript
<script runat="server" language="JScript">
Response.Write(2 ** 3); // Output: 8
var x = 3;
x **= 2;
Response.Write(x); // Output: 9
</script>
```

---

## BigInt Support

### Syntax

```javascript
var large = 100n;
var another = BigInt("9007199254740991");
```

### Remarks

- `BigInt` is a primitive wrapper object used to represent and manipulate primitive `bigint` values—which are too large to be represented by the `number` primitive.
- A `BigInt` value is created by appending `n` to the end of an integer literal, or by calling the `BigInt()` constructor.
- **Restriction:** You cannot mix `BigInt` and `Number` in the same operation (e.g., `10n + 5` throws `TypeError`). You must use explicit conversion.
- Arithmetic operations (`+`, `-`, `*`, `/`, `%`, `**`) and comparison operators are supported.
- `BigInt` division truncates towards zero.

### Code Example

```javascript
<script runat="server" language="JScript">
var a = 10n;
var b = 20n;
Response.Write(a + b); // Output: 30
Response.Write(2n ** 64n); // Output: 18446744073709551616

try {
    Response.Write(10n + 5);
} catch (e) {
    Response.Write("Error: " + e.message); // Output: Error: Cannot mix BigInt and other types...
}
</script>
```

---

## Promises

### Syntax

```javascript
var p = new Promise(function(resolve, reject) {
    // asynchronous operation
    if (success) resolve(data);
    else reject(error);
});

p.then(onFulfilled, onRejected)
 .catch(onRejected)
 .finally(onFinally);
```

### Remarks

- AxonASP implements the full ES6 `Promise` API.
- **Microtask Queue:** Promises are resolved using a Microtask queue. In the ASP environment, the queue is processed automatically when the script finishes or when an `await` is hit.
- Supported static methods: `Promise.resolve(v)`, `Promise.reject(r)`, `Promise.all(iterable)`, `Promise.race(iterable)`.

### Code Example

```javascript
<script runat="server" language="JScript">
var p = Promise.resolve(42);
p.then(function(val) {
    Response.Write("Promise resolved with: " + val);
});
</script>
```

---

## Generators (function*)

### Syntax

```javascript
function* myGenerator() {
    yield 1;
    yield 2;
    return 3;
}

var g = myGenerator();
var result = g.next(); // { value: 1, done: false }
```

### Remarks

- Generators are functions that can be exited and later re-entered. Their context (variable bindings) will be saved across re-entrances.
- Calling a generator function does not execute its body immediately; it returns an iterator object.
- `yield` pauses generator execution and returns a value to the caller.
- `yield*` delegates to another generator or iterable (currently implemented as basic yield).

### Code Example

```javascript
<script runat="server" language="JScript">
function* idMaker() {
    var index = 0;
    while (true)
        yield index++;
}

var gen = idMaker();
Response.Write(gen.next().value + "|"); // 0
Response.Write(gen.next().value + "|"); // 1
Response.Write(gen.next().value);       // 2
</script>
```

---

## Async/Await

### Syntax

```javascript
async function fetchData() {
    var response = await someAsyncOperation();
    return response.data;
}

fetchData().then(function(data) {
    Response.Write(data);
});
```

### Remarks

- `async` functions return a `Promise`.
- `await` pauses the execution of the async function until the promise is settled.
- **Synchronous Blocking:** In the AxonASP environment, `await` blocks the current request thread while pumping the microtask queue, ensuring predictable execution order for ASP pages.
- Standard `try...catch` blocks can be used to handle rejections from awaited promises.

### Code Example

```javascript
<script runat="server" language="JScript">
async function calculate(a, b) {
    var val = await Promise.resolve(a + b);
    return val * 2;
}

calculate(10, 5).then(function(result) {
    Response.Write("Result: " + result); // Output: Result: 30
});
</script>
```

---

## ECMAScript Modules (ESM)

### Syntax

```javascript
// math.js
export const PI = 3.14159;
export function add(a, b) { return a + b; }

// main.asp
import { PI, add } from './math.js';
Response.Write(add(PI, 10));
```

### Remarks

- AxonASP supports ES Modules via the `import` and `export` statements.
- **Global AST Cache:** Modules are read and compiled into AST/Bytecode ONCE globally and shared across all requests.
- **Request-Local Registry:** Each request has its own isolated module execution state. Top-level variables in a module are NOT shared between different users or subsequent requests.
- **Singleton per Request:** A module is executed only once within a single request, even if imported multiple times.
- **VM Reset:** Module instances are automatically cleared at the end of each request to prevent memory leaks and state contamination.
- **Module Resolution:** Imports are resolved relative to the current file path. Absolute paths and standard ASP virtual paths are also supported.

### Code Example

```javascript
<script runat="server" language="JScript">
// Assume 'config.js' exists with: export const version = "2.0";
import { version } from './config.js';
Response.Write("Application Version: " + version);
</script>
```

---

## Metaprogramming (Proxy and Reflect)

### Proxy

The `Proxy` object enables you to create a proxy for another object, which can intercept and redefine fundamental operations for that object.

#### Syntax

```javascript
var proxy = new Proxy(target, handler);
```

#### Remarks

- **Target:** The object to wrap (can be any object, including functions).
- **Handler:** An object whose properties are functions which define the behavior of the proxy when an operation is performed on it.
- Calling `Proxy` without `new` throws a `TypeError`.
- Both `target` and `handler` must be objects; otherwise, a `TypeError` is thrown.
- **Supported Traps:**
    - `get(target, property, receiver)`: Intercepts property reads.
    - `set(target, property, value, receiver)`: Intercepts property writes. Must return `true` to indicate success. In **Strict Mode**, returning `false` (or any falsy value) throws a `TypeError`.
    - `apply(target, thisArg, argumentsList)`: Intercepts function calls.
    - `construct(target, argumentsList, newTarget)`: Intercepts `new` operator calls. Must return an object.
- **Static Methods:**
    - `Proxy.revocable(target, handler)`: Creates a revocable `Proxy` object. Returns an object with two properties: `proxy` and `revoke`.
- **Support Note:** AxonASP now supports `get`, `set`, `apply`, and `construct` traps. Interception for other operations (like `has`, `deleteProperty`, `ownKeys`) will be implemented in subsequent phases.

### Reflect

`Reflect` is a built-in object that provides methods for interceptable JScript operations.

#### Remarks

- `Reflect` is a namespace object, not a constructor. You cannot use `new Reflect()`.
- **Support Note:** AxonASP currently provides the `Reflect` global namespace. Static methods (like `Reflect.get`, `Reflect.set`) will be implemented in subsequent phases.

### Code Example

```javascript
<script runat="server" language="JScript">
// 1. Basic Proxy creation
var target = { a: 1 };
var handler = {};
var proxy = new Proxy(target, handler);

Response.Write(typeof proxy); // Output: object
Response.Write(proxy instanceof Object); // Output: True

// 2. Reflect namespace
Response.Write(typeof Reflect); // Output: object
</script>
```


