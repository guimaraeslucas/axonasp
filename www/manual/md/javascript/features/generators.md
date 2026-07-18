# Generators (function*)

## Syntax

```javascript
function* myGenerator() {
    yield 1;
    yield 2;
    return 3;
}

var g = myGenerator();
var result = g.next(); // { value: 1, done: false }
```

## Remarks

- Generators are functions that can be exited and later re-entered. Their context (variable bindings) will be saved across re-entrances.
- Calling a generator function does not execute its body immediately; it returns an iterator object.
- `yield` pauses generator execution and returns a value to the caller.
- `yield*` delegates to another generator or iterable (currently implemented as basic yield).

## Code Example

```javascript
<script runat="server" language="JScript">
function* idMaker() {
    var index = 0;
    while (true)
        yield index++;
}

var gen = idMaker();
Response.Write(gen.next().value + "|"); // 0
Response.Write(gen.next().value + "|"); // 1
Response.Write(gen.next().value);       // 2
</script>
```
