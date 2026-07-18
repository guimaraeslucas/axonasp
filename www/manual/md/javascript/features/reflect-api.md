# Reflect API

## Syntax

```javascript
Reflect.get(target, propertyKey[, receiver])
Reflect.set(target, propertyKey, value[, receiver])
Reflect.has(target, propertyKey)
Reflect.deleteProperty(target, propertyKey)
Reflect.ownKeys(target)
Reflect.apply(target, thisArgument, argumentsList)
Reflect.construct(target, argumentsList[, newTarget])
```

## Remarks

- `Reflect` is a built-in object that provides methods for interceptable JScript operations.
- **Parity with Traps:** The methods are the same as those of proxy handlers.
- **Success Booleans:** Unlike standard operators, `Reflect.set` and `Reflect.deleteProperty` return a boolean indicating whether the operation succeeded, rather than throwing in strict mode.

## Code Example

```javascript
<script runat="server" language="JScript">
var obj = { x: 10 };
Reflect.set(obj, 'y', 20);
Response.Write(obj.y + "|" + Reflect.has(obj, 'x'));
// Output: 20|true

function greet(name) { return "Hello " + name; }
Response.Write(Reflect.apply(greet, undefined, ["World"]));
// Output: Hello World
</script>
```
