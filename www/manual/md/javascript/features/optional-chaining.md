# Optional Chaining (?.)

## Syntax

```javascript
obj?.property
obj?.[expression]
obj?.method()
```

## Remarks

- The optional chaining operator (`?.`) allows reading the value of a property located deep within a chain of connected objects without having to expressly validate that each reference in the chain is valid.
- If the object before the `?.` is `null` or `undefined`, the expression short-circuits and returns `undefined` instead of throwing an error.
- Works for property access, bracket access, and function calls.

## Code Example

```javascript
<script runat="server" language="JScript">
var user = { info: { name: "Alice" } };
Response.Write(user?.info?.name); // Output: Alice
Response.Write(user?.settings?.theme); // Output: undefined (no error)

var fn = null;
Response.Write(fn?.()); // Output: undefined (no error)
</script>
```
