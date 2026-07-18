# Default Parameter Values

## Syntax

```javascript
function greet(name, message = "Hello") {
    return message + ", " + name + "!";
}
```

## Remarks

- Native default parameter syntax is supported (for example, `function f(a = 10)`).
- The classic guard pattern `if (x === undefined) x = ...` is still supported and remains useful for compatibility-oriented scripts.

## Code Example

```javascript
<script runat="server" language="JScript">
function multiply(a, b = 2) {
    return a * b;
}
Response.Write(multiply(5));      // Output: 10
Response.Write(multiply(5, 3));   // Output: 15
</script>
```
