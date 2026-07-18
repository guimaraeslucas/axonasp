# Rest Parameters

## Syntax

```javascript
function fn(first, second, ...rest) {
    // rest is a standard array of remaining arguments
}
```

## Remarks

- The rest parameter must be the last parameter in the function signature.
- `rest` is a standard JScript array and supports all array methods.
- Only one rest parameter is allowed per function.

## Code Example

```javascript
<script runat="server" language="JScript">
function pack(head, ...rest) {
    return head + ":" + rest.length;
}
Response.Write(pack("h", 1, 2, 3));
// Output: h:3
</script>
```
