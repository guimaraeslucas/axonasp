# Property Reflection Helpers

## `Object.getOwnPropertyDescriptor(object, propertyName)`

Returns the property descriptor for an own property of `object`. The descriptor object contains the following fields: `value`, `writable`, `enumerable`, and `configurable`.

## `Object.getOwnPropertyDescriptors(object)`

Returns an object whose own properties are the property descriptors for all own properties of `object`. Each key maps to the same descriptor structure returned by `Object.getOwnPropertyDescriptor`.

## Remarks

- Both methods operate only on own properties. Inherited properties are not reported.
- Symbol-keyed internals follow the same visibility constraints as `Object.keys` and are not included in the result.
- `Object.defineProperty` is available and can be used to define non-enumerable or read-only properties before inspecting them with these helpers.

## Code Example

```javascript
<script runat="server" language="JScript">
var o = {};
Object.defineProperty(o, "hidden", {
    value: 10,
    writable: false,
    enumerable: false,
    configurable: false
});
var d = Object.getOwnPropertyDescriptor(o, "hidden");
var all = Object.getOwnPropertyDescriptors(o);
Response.Write(d.value + "|" + all.hidden.writable);
// Output: 10|false
</script>
```
