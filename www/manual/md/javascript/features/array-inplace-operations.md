# Array In-place Operations

## `Array.prototype.fill(value[, start[, end]])`

Fills all elements from `start` to `end` (exclusive) with `value`, in place. Negative indices are resolved relative to the array length. Returns the modified array.

## `Array.prototype.copyWithin(target[, start[, end]])`

Copies a portion of the array (from `start` to `end`, exclusive) to another position (`target`) within the same array, in place. Does not change the array length. Returns the modified array.

## `Array.prototype.keys()`

Returns an Array Iterator object containing each numeric index key from the array.

## `Array.prototype.entries()`

Returns an Array Iterator object containing `[index, value]` pairs for each array element.

## `Array.prototype.at(index)`

Returns the element at the specified `index`. Supports relative indexing from the end if `index` is negative.

## `Array.prototype.flat([depth])`

Returns a new array with all sub-array elements concatenated into it recursively up to the specified `depth`. Defaults to `1`.

## `Array.prototype.flatMap(callback[, thisArg])`

Returns a new array formed by applying a given callback function to each element of the array, and then flattening the result by one level.

## `Array.prototype.toSorted([compareFn])`

Returns a **new** array with the elements sorted in ascending order. Unlike `sort()`, it does not mutate the original array.

## `Array.prototype.toReversed()`

Returns a **new** array with the elements in reversed order. Unlike `reverse()`, it does not mutate the original array.

## `Array.prototype.toSpliced(start[, deleteCount[, ...items]])`

Returns a **new** array with some elements removed and/or replaced at a given index. Unlike `splice()`, it does not mutate the original array.

## Remarks

- Methods like `fill` and `copyWithin` operate in place and return the same array reference.
- `keys()` and `entries()` return standard iterable Array Iterator objects and can be consumed by `for...of`.
- Modern immutable methods (`toSorted`, `toReversed`, `toSpliced`) always return a new array instance.
- Negative index arguments in `at`, `fill`, and `copyWithin` are normalized relative to the array length.

## Code Example

```javascript
<script runat="server" language="JScript">
var arr = [1, [2, 3]];
Response.Write(JSON.stringify(arr.flat()));
// Output: [1,2,3]

var original = [3, 1, 2];
var sorted = original.toSorted();
Response.Write(sorted.join(","));
// Output: 1,2,3
Response.Write(original.join(","));
// Output: 3,1,2 (unchanged)

for (var k of [10, 20].keys()) {
    Response.Write(k + " ");
}
// Output: 0 1

for (var e of [10, 20].entries()) {
    Response.Write(e[0] + ":" + e[1] + " ");
}
// Output: 0:10 1:20

Response.Write("abc".at(-1));
// Output: c
</script>
```
