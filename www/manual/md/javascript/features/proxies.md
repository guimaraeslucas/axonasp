# Proxies

## Syntax

```javascript
var proxy = new Proxy(target, handler);
var { proxy, revoke } = Proxy.revocable(target, handler);
```

## Remarks

- The `Proxy` object allows you to create a proxy for another object, which can intercept and redefine fundamental operations for that object.
- **Supported traps:** `get`, `set`, `has`, `deleteProperty`, `apply`, `construct`, `ownKeys`, `defineProperty`, `getOwnPropertyDescriptor`, `getPrototypeOf`, `setPrototypeOf`, `isExtensible`, and `preventExtensions`.
- **Revocable Proxies:** `Proxy.revocable` returns an object with a `proxy` and a `revoke` function. Once revoked, any operation on the proxy throws a `TypeError`.
- **Centralized Trap Validator:** trap invariant checks are centralized in the runtime validator module to ensure consistent TypeError behavior across `Proxy` and `Reflect` paths.
- **Invariant Enforcement:** The engine strictly validates every ECMAScript section 10.5 trap invariant. A broken trap throws a `TypeError` with an `"invariant"` description. Key rules enforced:
    - **get:** A non-configurable, non-writable data property must return the exact stored value. A non-configurable accessor with no getter must return `undefined`.
  - **has:** Cannot return `false` for a non-configurable own property, or for any own property of a non-extensible target.
  - **set:** Cannot return `true` when the target has a non-configurable, non-writable data property and the new value differs from the stored one.
  - **deleteProperty:** Cannot return `true` for a non-configurable own property.
  - **ownKeys:** The result must include every non-configurable own property key. For non-extensible targets the result must exactly match the target's own key set.
  - **defineProperty:** Cannot return `true` when adding a new property to a non-extensible target, or when making a non-configurable property configurable.
  - **getOwnPropertyDescriptor:** Cannot return `undefined` for a non-configurable property. Cannot report a descriptor as `configurable: true` when the target property is non-configurable.
  - **getPrototypeOf:** For a non-extensible target must return the same object as the target's actual prototype.
  - **setPrototypeOf:** Cannot return `true` when the target is non-extensible and the supplied prototype differs from the target's current prototype.
  - **preventExtensions:** Cannot return `true` unless the target is already non-extensible at the time the trap returns.

## Code Example

```javascript
<script runat="server" language="JScript">
var target = { a: 1, b: 2 };
var handler = {
    get: function(target, prop, receiver) {
        if (prop === 'secret') return '***';
        return Reflect.get(target, prop, receiver);
    },
    has: function(target, prop) {
        if (prop === 'hidden') return true;
        return prop in target;
    }
};

var p = new Proxy(target, handler);
Response.Write(p.a + "|" + p.secret + "|" + ('hidden' in p));
// Output: 1|***|true
</script>
```
