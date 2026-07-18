# Array Search Utilities

## `Array.prototype.find(callback[, thisArg])`

Returns the first element that satisfies `callback`. Returns `undefined` when no element matches.

## `Array.prototype.findIndex(callback[, thisArg])`

Returns the index of the first element that satisfies `callback`. Returns `-1` when no element matches.

## Code Example

```javascript
<script runat="server" language="JScript">
var arr = [3, 7, 11, 14];
Response.Write(arr.find(function (x) { return x > 10; }));
// Output: 11
Response.Write(arr.findIndex(function (x) { return x > 10; }));
// Output: 2
</script>
```
