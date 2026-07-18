# Nullish Coalescing (??)

## Syntax

```javascript
var result = leftExpr ?? rightExpr;
```

## Remarks

- The nullish coalescing operator (`??`) is a logical operator that returns its right-hand side operand when its left-hand side operand is `null` or `undefined`, and otherwise returns its left-hand side operand.
- Unlike the OR operator (`||`), it does not return the right-hand side for other "falsy" values like `0`, `""`, or `false`.

## Code Example

```javascript
<script runat="server" language="JScript">
Response.Write(null ?? "default"); // Output: default
Response.Write(undefined ?? "default"); // Output: default
Response.Write(0 ?? 42); // Output: 0
Response.Write("" ?? "hello"); // Output: (empty string)
Response.Write(false ?? true); // Output: False
</script>
```
