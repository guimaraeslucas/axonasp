# Computed Property Names

## Syntax

```javascript
var key = "name";
var obj = { [key]: "Alice" };
var obj2 = { [prefix + "_en"]: "Hello", ["dynamic"]: 42 };
```

Use square brackets around a key expression inside an object literal to compute the property name at runtime.

## Remarks

- The expression inside `[...]` is evaluated at runtime and coerced to a string to form the property name.
- Any valid JScript expression can be used as the key: variables, string concatenations, function calls, and so on.
- Computed keys can be mixed freely with static keys and shorthand properties in the same literal.
- Numeric computed keys are coerced to strings before assignment (consistent with JScript's property model).

## Code Example

```javascript
<script runat="server" language="JScript">
var type = "color";
var o = {
    static: "fixed",
    [type]: "red",
    [type + "_code"]: "#FF0000"
};
Response.Write(o.static);       // Output: fixed
Response.Write(o.color);        // Output: red
Response.Write(o.color_code);   // Output: #FF0000

// Dynamic method name
var methodKey = "greet";
var api = { [methodKey]: function(n) { return "Hello, " + n; } };
Response.Write(api.greet("World")); // Output: Hello, World
</script>
```
