# ES6 Number Static Methods

The following static methods are available on the `Number` object.

## `Number.isInteger(value)`

Returns `true` only if `value` is a number with no fractional part and is not `Infinity` or `NaN`. Does **not** coerce non-number values; non-numbers return `false`.

## `Number.isNaN(value)`

Returns `true` only if `value` is the numeric `NaN`. Does **not** coerce non-number values; non-numbers always return `false`. This differs from the global `isNaN()` function, which coerces its argument.

## `Number.isFinite(value)`

Returns `true` only if `value` is a finite number. Does **not** coerce non-number values; non-numbers always return `false`.

## `Number.isSafeInteger(value)`

Returns `true` if `value` is an integer in the range `-(2^53 - 1)` to `2^53 - 1` inclusive, and has no fractional part. Does **not** coerce non-number values.

## `Number.parseInt(string, radix)`

Equivalent to the global `parseInt()` function. Parses `string` as an integer in the specified `radix` (2-36). Defaults to base 10.

## `Number.parseFloat(string)`

Equivalent to the global `parseFloat()` function. Parses `string` as a floating-point number.

## Number Constants

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

## Code Example

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
