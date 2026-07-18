# Atomics

## Syntax

```javascript
Atomics.add(typedArray, index, value)
Atomics.sub(typedArray, index, value)
Atomics.and(typedArray, index, value)
Atomics.or(typedArray, index, value)
Atomics.xor(typedArray, index, value)
Atomics.load(typedArray, index)
Atomics.store(typedArray, index, value)
Atomics.exchange(typedArray, index, value)
Atomics.compareExchange(typedArray, index, expectedValue, replacementValue)
Atomics.isLockFree(size)
```

## Remarks

- The `Atomics` object provides atomic operations as static methods. They are used with `SharedArrayBuffer` objects to ensure that concurrent memory accesses are predictable and safe.
- **Strict Validation:** In AxonASP, `Atomics` methods strictly require an integer TypedArray (e.g., `Int32Array`, `Uint8Array`) backed by a `SharedArrayBuffer`. Using a standard `ArrayBuffer` will throw a `TypeError`.
- **Atomic Operations:** These operations cannot be interrupted and are performed as a single unit. Even in the single-threaded context of a standard ASP request, they provide the necessary semantics for modern JavaScript libraries.
- `Atomics.isLockFree(size)` returns `true` for sizes 1, 2, 4, and 8, indicating that these operations are performed natively and efficiently by the CPU.

## SharedArrayBuffer

### Syntax

```javascript
var sab = new SharedArrayBuffer(byteLength);
```

### Remarks

- `SharedArrayBuffer` represents a generic, fixed-length raw binary data buffer, similar to `ArrayBuffer`.
- Unlike `ArrayBuffer`, a `SharedArrayBuffer` cannot be detached and its memory can be shared across multiple agents (workers).
- In the AxonASP single-threaded VM context, `SharedArrayBuffer` behaves identically to `ArrayBuffer` but provides the necessary API compatibility for modern libraries and prepares the engine for future multi-agent support.
- `SharedArrayBuffer` objects can be used as the backing store for any TypedArray or `DataView`.

### Code Example

```javascript
<script runat="server" language="JScript">
var sab = new SharedArrayBuffer(1024);
var u8 = new Uint8Array(sab);
u8[0] = 42;
Response.Write("Value: " + u8[0] + ", Length: " + sab.byteLength);
// Output: Value: 42, Length: 1024
</script>
```

## Code Example

```javascript
<script runat="server" language="JScript">
var sab = new SharedArrayBuffer(1024);
var u32 = new Uint32Array(sab);

Atomics.store(u32, 0, 100);
var old = Atomics.add(u32, 0, 50);

Response.Write("Old: " + old + ", New: " + Atomics.load(u32, 0));
// Output: Old: 100, New: 150
</script>
```
