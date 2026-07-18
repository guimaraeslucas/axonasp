# Weak References (WeakRef and FinalizationRegistry)

## Syntax

```javascript
var obj = { data: 42 };

// 1. WeakRef
var wr = new WeakRef(obj);
var target = wr.deref();
if (target !== undefined) {
    Response.Write(target.data);
}

// 2. FinalizationRegistry
var registry = new FinalizationRegistry(function(heldValue) {
    // Callback executed when registered objects are garbage collected
});
registry.register(obj, "some context", obj); // register object
registry.unregister(obj); // unregister
```

## Remarks

- **WeakRef:** Provides a way to hold a weak reference to an object, function, or symbol, allowing it to be garbage collected while still attempting to access it if it hasn't been collected yet via the `deref()` method.
- **FinalizationRegistry:** Allows you to register a callback to be invoked when an object is garbage collected.
- **VM Implementation Note:** AxonASP's JScript engine focuses on short-lived, high-performance HTTP request processing and does not implement a background garbage collector. Objects typically live until the end of the script execution (or request completion). Therefore, `FinalizationRegistry` callbacks will not be triggered during standard execution, but the API and validation semantics are fully implemented for compatibility with modern JavaScript libraries that expect these features to exist.
- Target objects for `WeakRef` and `FinalizationRegistry.register` must be Objects (`{}`), Functions, or unique Symbols (not registered via `Symbol.for()`). Passing primitives will result in a `TypeError`.

## Code Example

```javascript
<script runat="server" language="JScript">
var obj = { data: 42 };
var wr = new WeakRef(obj);
var target = wr.deref();
if (target !== undefined) {
    Response.Write(target.data); // Output: 42
}

var registry = new FinalizationRegistry(function(held) {
    // cleanup callback
});
registry.register(obj, "cleanup-tag");
Response.Write("WeakRef and FinalizationRegistry created successfully");
</script>
```
