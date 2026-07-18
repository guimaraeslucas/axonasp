# Full Unicode Support

## String Code Point Escapes

ES6 introduces a new escape sequence for Unicode characters that allows representing any character using its code point value in hexadecimal between braces.

## Syntax

```javascript
var s = "\u{1D306}"; // Tetragram for Centre
```

## Remarks

- Supports values from `0` up to `0x10FFFF`.
- Correctly handles surrogate pairs internally. A character like `\u{1D306}` has a `.length` of 2 in JScript (representing two UTF-16 code units).

## RegExp /u flag

The `u` flag (Unicode) enables advanced Unicode features in regular expressions.

### Syntax

```javascript
var re = /^\u{1D306}$/u;
```

### Remarks

- When the `u` flag is present, `.` matches a full Unicode code point (even if it spans multiple UTF-16 code units).
- Enables `\u{...}` escape sequences inside the regular expression pattern.

## Code Example

```javascript
<%
// String length with surrogate pairs
var s = "\u{1D306}";
Response.Write(s.length); // Output: 2

// Unicode RegExp matching
var re = /^.$/u;
Response.Write(re.test(s)); // Output: true (matches the whole surrogate pair)
%>
```
