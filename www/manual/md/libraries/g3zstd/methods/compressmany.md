# CompressMany Method

## Overview

Compresses an array of strings or byte arrays using the Zstandard (zstd) algorithm. This G3Pix AxonASP method allows for batch compression of multiple payloads efficiently.

## Syntax

```asp
resultArray = obj.CompressMany(inputArray [, level])
```

## Parameters and Arguments

- **inputArray**: A VBArray containing one or more elements. Each element must be a String or a VBArray of bytes.
- **level** (Optional): An Integer specifying the compression level. The range is -5 to 22. If omitted, the default level is used.

## Return Values

Returns a VBArray containing several VBArrays of bytes, where each sub-array corresponds to the compressed data of the input items.

## Remarks

- This method provides efficient batch processing for multiple payloads.
- If any input item fails validation or compression, an empty array or error may be returned depending on the context.
- The compression level remains constant for all items in the batch.
- Errors during batch processing are logged to the `LastError` property.

## Code Example

```asp
<%
Option Explicit
Dim objZstd, payloads, results, i
payloads = Array("Item 1", "Item 2", "Item 3")

Set objZstd = Server.CreateObject("G3ZSTD")

' Batch compress with default level
results = objZstd.CompressMany(payloads)

For i = 0 To UBound(results)
    Response.Write "Item " & i & " compressed size: " & UBound(results(i)) + 1 & " bytes.<br>"
Next

Set objZstd = Nothing
%>
```
