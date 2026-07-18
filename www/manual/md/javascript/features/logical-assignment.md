# Logical Assignment (||=, &&=, ??=)

## Syntax

```javascript
a ||= b;  // Logical OR assignment
a &&= b;  // Logical AND assignment
a ??= b;  // Nullish coalescing assignment
```

## Remarks

- `a ||= b` only assigns `b` to `a` if `a` is falsy.
- `a &&= b` only assigns `b` to `a` if `a` is truthy.
- `a ??= b` only assigns `b` to `a` if `a` is nullish (`null` or `undefined`).
- These operators short-circuit; the right-hand side is only evaluated if the assignment condition is met.

## Code Example

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
