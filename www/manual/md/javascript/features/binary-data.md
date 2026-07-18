# Binary Data - ArrayBuffer, SharedArrayBuffer and Typed Arrays

## Syntax

```javascript
var buffer = new ArrayBuffer(byteLength);
var sab    = new SharedArrayBuffer(byteLength);
var view   = new Uint8Array(buffer);
var view   = new Uint8Array(sab);
var view   = new Uint8Array(length);
var view   = new Uint8Array([1, 2, 3]);
var dv     = new DataView(buffer [, byteOffset [, byteLength]]);
var dv     = new DataView(sab [, byteOffset [, byteLength]]);
```

## Remarks

- `ArrayBuffer` holds a raw byte block. Its `byteLength` property returns its size in bytes. Use `ArrayBuffer.isView(v)` to test whether a value is a typed array view.
- `SharedArrayBuffer` is similar to `ArrayBuffer` but represents memory that can be shared between agents (workers). In the AxonASP single-threaded context, it behaves like a non-transferable `ArrayBuffer`.
- **Typed arrays** provide strongly-typed views over an `ArrayBuffer` or `SharedArrayBuffer`. All supported types are listed in the table below.
- `DataView` gives byte-level control over reads and writes including explicit endianness.
- Typed array constructors can be called with: a byte **length**, an existing **ArrayBuffer/SharedArrayBuffer**, or an **array-like** source (plain array or another typed array).
- Index reads past the end of the view return `undefined`. Index writes past the end are silently ignored.
- Calling a typed array constructor without `new` raises a `TypeError`.

## Supported Typed Array Types

| Constructor | Element type | Bytes per element |
|---|---|---|
| `Int8Array` | Signed 8-bit integer | 1 |
| `Uint8Array` | Unsigned 8-bit integer | 1 |
| `Uint8ClampedArray` | Unsigned 8-bit integer, clamped [0-255] | 1 |
| `Int16Array` | Signed 16-bit integer | 2 |
| `Uint16Array` | Unsigned 16-bit integer | 2 |
| `Int32Array` | Signed 32-bit integer | 4 |
| `Uint32Array` | Unsigned 32-bit integer | 4 |
| `Float32Array` | 32-bit IEEE 754 float | 4 |
| `Float64Array` | 64-bit IEEE 754 float | 8 |
| `BigInt64Array` | Signed 64-bit integer (BigInt) | 8 |
| `BigUint64Array` | Unsigned 64-bit integer (BigInt) | 8 |

## Typed Array Properties

| Property | Description |
|---|---|
| `length` | Number of elements |
| `byteLength` | Total size in bytes |
| `byteOffset` | Offset into the backing `ArrayBuffer` |
| `buffer` | The underlying `ArrayBuffer` |

## Typed Array Methods

| Method | Description |
|---|---|
| `set(array [, offset])` | Copy values from an array-like source |
| `subarray([begin [, end]])` | Return a new view over the same buffer |
| `fill(value [, start [, end]])` | Fill all or part of the view with a value |
| `slice([begin [, end]])` | Return a new typed array copy of the range |
| `forEach(callback)` | Iterate over each element |
| `indexOf(value [, fromIndex])` | Return first index of a matching value, or -1 |

## ArrayBuffer Methods

| Method | Description |
|---|---|
| `slice([begin [, end]])` | Return a new `ArrayBuffer` containing a copy of the byte range |
| `ArrayBuffer.isView(value)` | Return `true` if the value is a typed array or DataView |

## DataView Methods

| Method | Description |
|---|---|
| `getInt8(offset)` | Read signed 8-bit int |
| `getUint8(offset)` | Read unsigned 8-bit int |
| `getInt16(offset [, littleEndian])` | Read signed 16-bit int |
| `getUint16(offset [, littleEndian])` | Read unsigned 16-bit int |
| `getInt32(offset [, littleEndian])` | Read signed 32-bit int |
| `getUint32(offset [, littleEndian])` | Read unsigned 32-bit int |
| `getFloat32(offset [, littleEndian])` | Read 32-bit float |
| `getFloat64(offset [, littleEndian])` | Read 64-bit float |
| `setInt8(offset, value)` | Write signed 8-bit int |
| `setUint8(offset, value)` | Write unsigned 8-bit int |
| `setInt16(offset, value [, littleEndian])` | Write signed 16-bit int |
| `setUint16(offset, value [, littleEndian])` | Write unsigned 16-bit int |
| `setInt32(offset, value [, littleEndian])` | Write signed 32-bit int |
| `setUint32(offset, value [, littleEndian])` | Write unsigned 32-bit int |
| `setFloat32(offset, value [, littleEndian])` | Write 32-bit float |
| `setFloat64(offset, value [, littleEndian])` | Write 64-bit float |

## Code Example

```javascript
<script runat="server" language="JScript">
// --- ArrayBuffer and Uint8Array ---
var buffer = new ArrayBuffer(4);
var view = new Uint8Array(buffer);
view[0] = 10;
view[1] = 20;
view[2] = 30;
view[3] = 40;
Response.Write(view[0] + "," + view[1] + "," + view[2] + "," + view[3]);
// Output: 10,20,30,40

// --- Uint8ClampedArray ---
var clamped = new Uint8ClampedArray(2);
clamped[0] = 300; // clamped to 255
clamped[1] = -5;  // clamped to 0
Response.Write(clamped[0] + "," + clamped[1]);
// Output: 255,0

// --- Int32Array from plain array ---
var ints = new Int32Array([-100, 0, 100]);
Response.Write(ints[0] + "," + ints.byteLength);
// Output: -100,12

// --- DataView with explicit endianness ---
var db = new ArrayBuffer(8);
var dv = new DataView(db);
dv.setInt32(0, 0xDEADBEEF, false); // big-endian
Response.Write(dv.getInt32(0, false));
// Output: -559038737

// --- ArrayBuffer.slice ---
var sliced = buffer.slice(1, 3);
Response.Write(sliced.byteLength);
// Output: 2

// --- for...of on typed array ---
var a = new Uint8Array([10, 20, 30]);
var sum = 0;
for (var v of a) { sum += v; }
Response.Write(sum);
// Output: 60
</script>
```
