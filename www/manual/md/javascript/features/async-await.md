# Async/Await

## Syntax

```javascript
async function fetchData() {
    var response = await someAsyncOperation();
    return response.data;
}

fetchData().then(function(data) {
    Response.Write(data);
});
```

## Remarks

- `async` functions return a `Promise`.
- `await` pauses the execution of the async function until the promise is settled.
- **Synchronous Blocking:** In the AxonASP environment, `await` blocks the current request thread while pumping the microtask queue, ensuring predictable execution order for ASP pages.
- Standard `try...catch` blocks can be used to handle rejections from awaited promises.

## Code Example

```javascript
<script runat="server" language="JScript">
async function calculate(a, b) {
    var val = await Promise.resolve(a + b);
    return val * 2;
}

calculate(10, 5).then(function(result) {
    Response.Write("Result: " + result); // Output: Result: 30
});
</script>
```
