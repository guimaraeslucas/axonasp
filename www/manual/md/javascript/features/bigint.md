# BigInt Support

## Syntax

```javascript
var large = 100n;
var another = BigInt("9007199254740991");
```

## Remarks

- `BigInt` is a primitive wrapper object used to represent and manipulate primitive `bigint` values which are too large to be represented by the `number` primitive.
- A `BigInt` value is created by appending `n` to the end of an integer literal, or by calling the `BigInt()` constructor.
- **Restriction:** You cannot mix `BigInt` and `Number` in the same operation (e.g., `10n + 5` throws `TypeError`). You must use explicit conversion.
- Arithmetic operations (`+`, `-`, `*`, `/`, `%`, `**`) and comparison operators are supported.
- `BigInt` division truncates towards zero.

## Code Example

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
