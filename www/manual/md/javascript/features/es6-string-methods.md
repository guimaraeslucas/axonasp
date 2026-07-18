# ES6 String Methods

The following methods are available on `String` values.

## `String.prototype.includes(searchString[, position])`

Returns `true` if `searchString` is found anywhere within the string at or after `position` (default `0`); `false` otherwise. Case-sensitive. Raises a `TypeError` if `searchString` is a `RegExp`.

## `String.prototype.startsWith(searchString)`

Returns `true` if the string begins with `searchString`; `false` otherwise. Case-sensitive.

## `String.prototype.endsWith(searchString)`

Returns `true` if the string ends with `searchString`; `false` otherwise. Case-sensitive.

## `String.prototype.repeat(count)`

Returns a new string containing `count` repetitions of the original string. Returns an empty string if `count` is 0.

## `String.prototype.at(index)`

Returns the character at the specified `index`. Supports relative indexing from the end if `index` is negative.

## `String.prototype.codePointAt(position)`

Returns the Unicode code point value at `position`. If `position` is out of range, returns `undefined`.

## `String.prototype.normalize([form])`

Returns the Unicode Normalization Form of the string. Supported values are `NFC`, `NFD`, `NFKC`, and `NFKD`. If omitted, `NFC` is used.

## `String.prototype.padStart(targetLength, padString)`

Pads the string from the start with `padString` until the total length reaches `targetLength`. If `padString` is not supplied, spaces are used.

## `String.prototype.padEnd(targetLength, padString)`

Pads the string from the end with `padString` until the total length reaches `targetLength`. If `padString` is not supplied, spaces are used.

## Code Example

```javascript
<script runat="server" language="JScript">
var s = "Hello World";

Response.Write(s.includes("World"));        // Output: true
Response.Write(s.includes("World", 6));     // Output: true
Response.Write(s.startsWith("Hello"));      // Output: true
Response.Write(s.endsWith("World"));        // Output: true
Response.Write("ab".repeat(3));             // Output: ababab
Response.Write("5".padStart(3, "0"));       // Output: 005
Response.Write("5".padEnd(3, "0"));         // Output: 500
Response.Write("A😀B".codePointAt(1));       // Output: 128512
Response.Write("e\u0301".normalize("NFC")); // Output: é

var regexError = false;
try {
    "hello".includes(new RegExp("h"));
} catch (e) {
    regexError = String(e).indexOf("TypeError") !== -1;
}
Response.Write(regexError);                 // Output: true
</script>
```
