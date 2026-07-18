# Template Literals

## Syntax

```javascript
var result = `static text ${expression} more static text`;
```

Template literals are enclosed in backticks (`` ` ``). They support embedded expressions using `${expression}` placeholders and preserve literal newlines.

## Remarks

- All `${expression}` placeholders are evaluated at runtime and coerced to strings using standard JScript string coercion.
- Multiple expressions can be embedded in a single template literal.
- Multi-line template literals preserve embedded newline characters.
- Tagged template literals are not supported. A tagged template (e.g., `` tag`...` ``) resolves to `undefined`.

## Code Example

```javascript
<%
var name = "World";
var count = 42;
var msg = `Hello, ${name}! You have ${count} messages.`;
Response.Write(msg);
// Output: Hello, World! You have 42 messages.

var a = 3, b = 4;
Response.Write(`Sum: ${a + b}`);
// Output: Sum: 7
%>
```
