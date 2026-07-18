# Spread in Array Literals

## Syntax

```javascript
var out = [1, 2, ...otherArray, 5];
```

## Remarks

- Spread in array literals expands one source array-like value into individual elements.
- `null` and `undefined` spread sources raise a JScript `TypeError`.
- Evaluation order is preserved left to right.

## Code Example

```javascript
<script runat="server" language="JScript">
var src = [3, 4];
var out = [1, 2, ...src, 5];
Response.Write(out.join(","));
// Output: 1,2,3,4,5
</script>
```
