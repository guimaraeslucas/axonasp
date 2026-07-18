# Weak Collections (WeakMap and WeakSet)

## Syntax

```javascript
var wm = new WeakMap();
var ws = new WeakSet();

var key = {};
wm.set(key, "data");
ws.add(key);

Response.Write(wm.get(key)); // data
Response.Write(ws.has(key)); // True
```

## Remarks

- `WeakMap` and `WeakSet` provide collections where keys (or values in `WeakSet`) are held weakly.
- **Memory Safety:** Unlike standard `Map` and `Set`, weak collections do not prevent their keys from being garbage collected. This is critical for preventing memory leaks in long-running scripts where objects are used as temporary keys.
- **Inverted Storage:** AxonASP uses an efficient "inverted storage" pattern where weak data is stored internally within the key object itself, ensuring that when the key is destroyed, the associated data is automatically reclaimed without GC overhead.
- **Valid Keys:** Objects (`{}`), functions (`function`), and **unique Symbols** (those created via `Symbol()` that are not registered in the global registry via `Symbol.for()` and are not well-known symbols like `Symbol.iterator`) can be used as keys. Attempting to use a primitive (string, number, boolean) or a restricted symbol as a key will throw a `TypeError`.
- **Non-Iterable:** Weak collections are not iterable. They do not support `for...of` loops, and they do not have `.size`, `.keys()`, `.values()`, or `.entries()` methods.

## Code Example

```javascript
<script runat="server" language="JScript">
var obj = {};
var wm = new WeakMap();
wm.set(obj, "confidential");
Response.Write(wm.get(obj)); // Output: confidential

var ws = new WeakSet();
ws.add(obj);
Response.Write(ws.has(obj)); // Output: true
</script>
```
