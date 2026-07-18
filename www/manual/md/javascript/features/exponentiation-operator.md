# Exponentiation Operator (**)

## Syntax

```javascript
var result = base ** exponent;
var a **= exponent;
```

## Remarks

- The exponentiation operator (`**`) returns the result of raising the first operand to the power of the second operand.
- It is equivalent to `Math.pow()`, but also supports `BigInt`.

## Code Example

```javascript
<script runat="server" language="JScript">
Response.Write(2 ** 3); // Output: 8
var x = 3;
x **= 2;
Response.Write(x); // Output: 9
</script>
```
