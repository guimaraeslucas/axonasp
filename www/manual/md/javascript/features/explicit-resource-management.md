# Explicit Resource Management (using)

## Syntax

```javascript
using resource = expression;
async using asyncResource = expression;
```

## Remarks

- `using` is supported as a block-scoped declaration and requires an initializer.
- Resource cleanup uses symbol-based disposal methods:
  - `Symbol.dispose` for `using`
  - `Symbol.asyncDispose` for `async using`
- Multiple `using` declarations in the same scope are disposed in reverse declaration order.
- Disposal runs at normal scope exit and during exception unwinding (`throw`) for the same scope.
- `async using` currently invokes `Symbol.asyncDispose` synchronously (without awaiting Promise settlement).

## Code Example

```javascript
<script runat="server" language="JScript">
var trace = [];

var firstResource = {
    [Symbol.dispose]: function() {
        trace.push("dispose:first");
    }
};

var secondResource = {
    [Symbol.dispose]: function() {
        trace.push("dispose:second");
    }
};

{
    using first = firstResource;
    using second = secondResource;
    trace.push("inside");
}

Response.Write(trace.join("|"));
// Output: inside|dispose:second|dispose:first
</script>
```
