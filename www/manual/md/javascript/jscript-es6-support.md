# Use ES6 Features in Javascript Scripts

## Overview

AxonASP's Javascript engine supports a wide range of modern ECMAScript features, including ES6 (ES2015) additions and subsequent standards up to ES2024. This page documents all supported modern capabilities: template literals, block-scoped declarations (`let` and `const`) with Temporal Dead Zone (TDZ), arrow functions, default parameter values, rest parameters, spread in array literals, object literal shorthand, computed property names, `for...of` loops, `Object` static utilities (including `values`, `entries`, and `fromEntries`), property reflection helpers, modern `String` methods (like `includes`, `padStart`, and `at`), full Unicode support in `RegExp`, `Number` static methods, `Math` extensions, `Symbol` primitive, `Set` and `Map` collections, and a comprehensive set of `Array` utilities (including `find`, `flat`, `flatMap`, and immutable `toSorted`/`toReversed`/`toSpliced` methods).

All ES6 features described here are available in `<script runat="server" language="JScript">` blocks and in `<% language="JScript" %>` inline blocks.

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

### Remarks

- `Object.assign` skips `null` and `undefined` sources.
- `Object.keys`, `Object.values`, and `Object.entries` throw a JScript `TypeError` when called with `null` or `undefined`.
- Return values are standard JScript arrays and are compatible with existing array operations.
- Symbol-keyed properties are intentionally excluded from `Object.keys`, `Object.values`, and `Object.entries` to reduce collision risks in legacy code.

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

## Math Extensions

The following additional methods are available on the `Math` object.

### `Math.trunc(x)`

Returns the integer part of `x` by removing the fractional digits.

### `Math.sign(x)`

Returns `1` for positive values, `-1` for negative values, and `0` for zero. Returns `NaN` for `NaN` input.

### `Math.cbrt(x)`

Returns the cube root of `x`.

### Code Example

```javascript
<script runat="server" language="JScript">
Response.Write(Math.trunc(4.9)); // Output: 4
Response.Write(Math.sign(-12));  // Output: -1
Response.Write(Math.cbrt(27));   // Output: 3
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

## For...Of Loops

### Syntax

```javascript
for (var element of iterable) { /* body */ }
for (let element of iterable) { /* body */ }
for (const element of iterable) { /* body */ }
```

The `for...of` loop iterates over the **values** of an iterable object in sequence.

### Supported Iterables

| Type | Behavior |
|---|---|
| Array (JS) | Yields each element by index in order. |
| String | Yields each character as a single-character string. |
| Set | Yields each unique member. |
| Map | Yields each `[key, value]` pair as a two-element array. |

### Remarks

- `var`, `let`, and `const` declarations are all supported in the loop header.
- `break` exits the loop immediately and discards the iterator.
- `continue` advances to the next value without executing the rest of the body.
- Nested `for...of` loops are supported.
- Iterating over an empty array or an empty string executes the body zero times.
- The source iterable is evaluated once before the first iteration; mutations to the source after the loop starts do not affect the values being iterated.

### Code Example

```javascript
<script runat="server" language="JScript">
// Array
var total = 0;
for (var n of [10, 20, 30]) {
    total += n;
}
Response.Write(total);
// Output: 60

// String
var chars = "";
for (var ch of "Hello") {
    chars += ch + "-";
}
Response.Write(chars);
// Output: H-e-l-l-o-

// Break
var found = false;
for (let x of [1, 2, 3, 4, 5]) {
    if (x === 3) { found = true; break; }
}
Response.Write(found);
// Output: true

// Set
var s = new Set();
s.add("a"); s.add("b"); s.add("c");
var setResult = "";
for (const v of s) {
    setResult += v;
}
Response.Write(setResult.length);
// Output: 3 (order may vary)
</script>
```
