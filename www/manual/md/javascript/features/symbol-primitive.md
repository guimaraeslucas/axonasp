# Symbol Primitive

## Syntax

```javascript
var sym = Symbol(description);
```

## Remarks

- Each call to `Symbol()` returns a unique value that is never equal to any other `Symbol` or primitive.
- Symbols can be used as object property keys to create collision-safe identifiers.
- Calling `new Symbol()` raises a `TypeError`. `Symbol` is not a constructor.
- Symbol-keyed properties are intentionally hidden from `Object.keys`, `Object.values`, and `Object.entries` to prevent unintended exposure in enumeration.

## Code Example

```javascript
<script runat="server" language="JScript">
var s1 = Symbol("id");
var s2 = Symbol("id");
var o = {};
o[s1] = 42;
Response.Write((s1 !== s2) + "|" + o[s1]);
// Output: true|42
</script>
```
