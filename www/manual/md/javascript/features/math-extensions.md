# Math Extensions

The following additional methods are available on the `Math` object.

## `Math.trunc(x)`

Returns the integer part of `x` by removing the fractional digits.

## `Math.sign(x)`

Returns `1` for positive values, `-1` for negative values, and `0` for zero. Returns `NaN` for `NaN` input.

## `Math.cbrt(x)`

Returns the cube root of `x`.

## Additional Methods

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

## Code Example

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
