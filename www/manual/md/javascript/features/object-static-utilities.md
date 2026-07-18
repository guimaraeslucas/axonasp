# Object Static Utilities

The following `Object` static methods are available.

## `Object.assign(target, ...sources)`

Copies enumerable own properties from each source object into `target`, from left to right, and returns `target`.

## `Object.keys(object)`

Returns an array of enumerable own property names.

## `Object.values(object)`

Returns an array of enumerable own property values.

## `Object.entries(object)`

Returns an array where each item is a two-element `[key, value]` pair for each enumerable own property.

## `Object.fromEntries(iterable)`

Converts an iterable of key-value pairs (such as an array of `[key, value]` arrays) into a new object.

## `Object.is(value1, value2)`

Returns `true` when both values are the same according to ECMAScript `Object.is` semantics. `NaN` compares equal to `NaN`, and `+0` and `-0` compare as different values.

## `Object.setPrototypeOf(object, prototype)`

Changes the prototype of `object` to `prototype` when the object is extensible. Throws a `TypeError` if the target is not an object or if its prototype cannot be changed.

## `Object.getOwnPropertySymbols(object)`

Returns an array of the object's own symbol-keyed properties in symbol form.

## Remarks

- `Object.assign` skips `null` and `undefined` sources.
- `Object.keys`, `Object.values`, and `Object.entries` throw a JScript `TypeError` when called with `null` or `undefined`.
- Return values are standard JScript arrays and are compatible with existing array operations.
- Symbol-keyed properties are intentionally excluded from `Object.keys`, `Object.values`, and `Object.entries` to reduce collision risks in legacy code.
- `Object.getOwnPropertySymbols` ignores string-keyed properties. Prototype inheritance is not included; only own properties are reported.

## Code Example

```javascript
<script runat="server" language="JScript">
var target = { a: 1 };
Object.assign(target, { b: 2 }, { c: 3 });

Response.Write(Object.keys(target).join(","));
// Output: a,b,c

Response.Write(Object.values(target).join(","));
// Output: 1,2,3

var e = Object.entries(target);
Response.Write(e[0][0] + ":" + e[0][1]);
// Output: a:1

var entries = [["x", 10], ["y", 20]];
var obj = Object.fromEntries(entries);
Response.Write(obj.x + "," + obj.y);
// Output: 10,20

var s = Symbol("id");
var o = {};
Object.setPrototypeOf(o, { inherited: 1 });
o[s] = 99;
Response.Write(Object.is(NaN, NaN));
// Output: true
Response.Write(Object.getOwnPropertySymbols(o).length);
// Output: 1
</script>
```
