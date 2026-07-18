# Global URI Functions

The following URI helper functions are available globally.

## `encodeURI(uri)`

Encodes a complete URI string while preserving URI-reserved separators such as `:`, `/`, `?`, `&`, `=`, and `#`.

## `decodeURI(uri)`

Decodes a complete URI string. Reserved separators remain preserved when they were percent-encoded.

## `encodeURIComponent(component)`

Encodes a URI component (such as one query value) and escapes reserved characters like `=`, `&`, and `+`.

## `decodeURIComponent(component)`

Decodes an encoded URI component.

## Code Example

```javascript
<script runat="server" language="JScript">
var full = "https://example.com/a path/?q=hello world&x=1+2#frag";
Response.Write(encodeURI(full));
// Output: https://example.com/a%20path/?q=hello%20world&x=1+2#frag

var component = "q=hello world&x=1+2";
var encoded = encodeURIComponent(component);
Response.Write(encoded);
// Output: q%3Dhello%20world%26x%3D1%2B2

Response.Write(decodeURIComponent(encoded));
// Output: q=hello world&x=1+2
</script>
```
