# new.target meta-property

## Syntax

```javascript
new.target
```

## Remarks

- `new.target` allows you to detect whether a function or constructor was called using the `new` operator.
- In constructors and functions invoked via the `new` operator, `new.target` returns a reference to the constructor or function.
- In normal function calls, `new.target` is `undefined`.
- This is particularly useful in class constructors to identify the specific class being instantiated (especially in inheritance scenarios).

## Code Example

```javascript
<script runat="server" language="JScript">
function Foo() {
  if (!new.target) {
    Response.Write("Called as function");
  } else {
    Response.Write("Called with new");
  }
}

Foo();      // Output: Called as function
new Foo();  // Output: Called with new
</script>
```
