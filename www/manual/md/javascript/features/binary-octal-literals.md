# Binary and Octal Numeric Literals

## Syntax

```javascript
var b = 0b1010; // binary
var o = 0o744;  // octal
```

## Remarks

- Prefix `0b` or `0B` parses base-2 integer literals.
- Prefix `0o` or `0O` parses base-8 integer literals.

## Code Example

```javascript
<script runat="server" language="JScript">
Response.Write(0b1010); // Output: 10
Response.Write(0o744);  // Output: 484
</script>
```
