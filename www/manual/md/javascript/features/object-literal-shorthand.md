# Object Literal Property Shorthand

## Syntax

```javascript
var x = 10;
var y = 20;
var point = { x, y }; // equivalent to { x: x, y: y }
```

## Remarks

- Shorthand property syntax is supported when the variable name and the property name are identical.
- Method shorthand (e.g., `{ greet() {} }`) follows the same rule and is available as well.

## Code Example

```javascript
<script runat="server" language="JScript">
var x = 10;
var y = 20;
var p = { x, y };
Response.Write(p.x + "," + p.y);
// Output: 10,20
</script>
```
