# Symbol Primitive - Well-Known Symbols and Global Registry

Well-known symbols are pre-defined `Symbol` values stored as properties of the `Symbol` constructor object.

| Symbol | Description |
|---|---|
| `Symbol.iterator` | Default iterator for `for...of` loops |
| `Symbol.toStringTag` | Object `[object X]` tag override |
| `Symbol.species` | Species constructor for derived objects |
| `Symbol.hasInstance` | Custom `instanceof` behavior |
| `Symbol.toPrimitive` | Custom primitive conversion |

The global symbol registry allows sharing symbols across realms via `Symbol.for` and `Symbol.keyFor`.

## Code Example

```javascript
<script runat="server" language="JScript">
// Well-known symbols are of type "symbol"
Response.Write(typeof Symbol.iterator);   // Output: symbol
Response.Write(typeof Symbol.toStringTag); // Output: symbol

// Symbol.for - global registry: same key returns same symbol
var a = Symbol.for("appToken");
var b = Symbol.for("appToken");
Response.Write(a === b); // Output: true

// Symbol.keyFor - retrieve key from registry
Response.Write(Symbol.keyFor(a)); // Output: appToken

// Locally created symbols are NOT in the registry
var local = Symbol("local");
Response.Write(Symbol.keyFor(local) === undefined); // Output: true
</script>
```
