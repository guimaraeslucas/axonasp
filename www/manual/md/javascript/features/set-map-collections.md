# Set and Map Collections

## `Set`

A `Set` stores unique values. Duplicate values are silently ignored on insertion.

| Method | Description |
|---|---|
| `set.add(value)` | Inserts `value` and returns the `Set`. |
| `set.has(value)` | Returns `true` if `value` is present. |
| `set.delete(value)` | Removes `value`. Returns `true` if the value existed. |
| `set.clear()` | Removes all elements. |
| `set.size` | Returns the number of unique elements. |

## `Map`

A `Map` stores key/value pairs and preserves insertion order.

| Method | Description |
|---|---|
| `map.set(key, value)` | Sets the entry for `key` and returns the `Map`. |
| `map.get(key)` | Returns the value associated with `key`, or `undefined`. |
| `map.has(key)` | Returns `true` if an entry for `key` exists. |
| `map.delete(key)` | Removes the entry for `key`. Returns `true` if it existed. |
| `map.clear()` | Removes all entries. |
| `map.size` | Returns the number of entries. |

## Code Example

```javascript
<script runat="server" language="JScript">
var s = new Set();
s.add("a");
s.add("b");
s.add("a"); // duplicate, ignored
Response.Write(s.has("a") + "|" + s.size);
// Output: true|2

var m = new Map();
m.set("k", 10);
Response.Write(m.has("k") + "|" + m.get("k"));
// Output: true|10
</script>
```
