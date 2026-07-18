# Block-Scoped Declarations (let and const)

## Syntax

```javascript
let x = 10;
const y = 20;

{
    let x = 30; // Shadows outer x
    const y = 40; // Shadows outer y
}
```

## Remarks

- `let` and `const` provide block-level scoping. Variables declared inside a `{}` block are only accessible within that block.
- **Temporal Dead Zone (TDZ):** Unlike `var`, accessing a `let` or `const` variable before its declaration line in the execution flow results in a `ReferenceError`.
- `const` bindings are immutable; attempting to reassign a value to a `const` variable results in a `TypeError`.

## Code Example

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
