# Arrow Functions

## Syntax

```javascript
// Concise body (expression result is implicitly returned)
var fn = (param1, param2) => expression;

// Block body
var fn = (param1, param2) => {
    // statements
    return value;
};
```

## Remarks

- Arrow functions do not create their own `this` binding. The value of `this` is captured **lexically** from the enclosing scope at the time the arrow function is created. This is useful for callbacks inside constructor methods.
- Arrow functions cannot be used as constructors. Using `new` with an arrow function is not supported.
- Single-parameter arrow functions without parentheses (e.g., `x => x * 2`) are supported.
- Arrow functions have an `arguments` object bound to the enclosing function's `arguments`, not their own.

## Code Example

```javascript
<script runat="server" language="JScript">
// Concise arrow function
var square = (x) => x * x;
Response.Write(square(5));
// Output: 25

// Lexical this in a constructor
function Timer() {
    this.seconds = 0;
    this.tick = function() {
        var increment = () => { this.seconds = this.seconds + 1; };
        increment();
    };
}
var t = new Timer();
t.tick();
t.tick();
Response.Write(t.seconds);
// Output: 2
</script>
```
