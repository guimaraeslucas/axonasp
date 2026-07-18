# Destructuring Assignment

Destructuring assignment is a syntax that makes it possible to unpack values from arrays, or properties from objects, into distinct variables.

## Object Destructuring

Object destructuring allows you to extract multiple properties from an object and assign them to variables in a single statement.

### Syntax

```javascript
var { p1, p2 } = object;
var { p1: v1, p2: v2 } = object;
var { p1 = defaultValue } = object; // Default value
var { p1, ...rest } = object; // Object rest
```

### Remarks

- If a variable name matches a property name, you can use the shorthand `{ p1, p2 }`.
- You can map a property to a different variable name using `{ property: variable }`.
- **Default Values:** You can provide a default value using `=`. The default is only used if the property is strictly `undefined`.
- **Rest Property:** The `...rest` syntax creates a new object containing all remaining enumerable own properties of the source object.
- Nested destructuring is supported: `var { a: { b } } = obj;`.
- Computed property names can be used: `var { [key]: value } = obj;`.
- **Validation:** Attempting to destructure `null` or `undefined` raises a `TypeError`.

### Code Example

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

## Array Destructuring

Array destructuring allows you to extract elements from arrays or any iterable object (like Strings, Sets, or Maps) using an array-like syntax.

### Syntax

```javascript
var [a, b] = iterable;
var [a, , c] = iterable; // Elision (skipping elements)
var [a = 10] = iterable; // Default value
var [a, ...rest] = iterable; // Array rest
var [a, [b, c]] = iterable; // Nested array destructuring
```

### Remarks

- Values are extracted in order from the source iterable.
- **Elision:** You can skip elements using commas: `var [first, , last] = [1, 2, 3];`.
- **Default Values:** Assigns a default if the iterable yields `undefined` or is exhausted.
- **Rest Elements:** The `...rest` syntax collects all remaining yielded values into a new Array.
- **Iteration Protocol:** Unlike object destructuring, array destructuring works with any object that implements the ES6 Iteration Protocol. This includes Strings (yielding characters), Maps (yielding `[key, value]` pairs), and Sets.
- **Validation:** Attempting to destructure a non-iterable value (like `true` or a plain object without `[Symbol.iterator]`) raises a `TypeError`.

### Code Example

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
