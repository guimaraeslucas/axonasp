# DecompressMany Method

## Overview

Decompresses an array of Zstandard (zstd) compressed payloads (strings or byte arrays). This G3Pix AxonASP method allows for efficient batch decompression of multiple datasets.

## Syntax

```asp
resultArray = obj.DecompressMany(inputArray)
```

## Parameters and Arguments

- **inputArray**: A VBArray containing one or more elements. Each element must be a String or a VBArray of bytes containing a Zstandard compressed frame.

## Return Values

Returns a VBArray containing several VBArrays of bytes, where each sub-array corresponds to the decompressed data of the input items.

## Remarks

- Each item in the input array must be a valid Zstandard frame.
- The method processes items in sequence and returns an array of the same length as the input.
- If any item fails decompression, the overall operation may result in an empty array or raise a runtime error.
- Errors during batch processing are logged to the `LastError` property.

## Code Example

```asp
<%
Option Explicit
Dim objZstd, compressedItems, decompressedItems, i
' Assuming compressedItems is an array of compressed data

Set objZstd = Server.CreateObject("G3ZSTD")

' Batch decompress
decompressedItems = objZstd.DecompressMany(compressedItems)

For i = 0 To UBound(decompressedItems)
    Response.Write "Item " & i & " decompressed size: " & UBound(decompressedItems(i)) + 1 & " bytes.<br>"
Next

Set objZstd = Nothing
%>
```
