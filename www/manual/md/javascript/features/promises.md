# Promises

## Syntax

```javascript
var p = new Promise(function(resolve, reject) {
    // asynchronous operation
    if (success) resolve(data);
    else reject(error);
});

p.then(onFulfilled, onRejected)
 .catch(onRejected)
 .finally(onFinally);
```

## Remarks

- AxonASP implements the full ES6 `Promise` API.
- **Microtask Queue:** Promises are resolved using a Microtask queue. In the ASP environment, the queue is processed automatically when the script finishes or when an `await` is hit.
- Supported static methods: `Promise.resolve(v)`, `Promise.reject(r)`, `Promise.all(iterable)`, `Promise.race(iterable)`.

## Code Example

```javascript
<script runat="server" language="JScript">
var p = Promise.resolve(42);
p.then(function(val) {
    Response.Write("Promise resolved with: " + val);
});
</script>
```
