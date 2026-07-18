# Array Construction Utilities

## `Array.from(arrayLike[, mapFn])`

Converts an array-like or iterable object into a standard JScript array. Accepts an optional mapping function that is applied to each element.

## `Array.of(...items)`

Creates a new array from its arguments. Unlike `new Array(n)`, `Array.of(n)` always creates a one-element array containing `n`.

## Code Example

```javascript
<script runat="server" language="JScript">
var a = Array.from({ length: 2, 0: "x", 1: "y" });
var b = Array.of(7, 8, 9);
Response.Write(a.join("-") + "|" + b.join("-"));
// Output: x-y|7-8-9
</script>
```
