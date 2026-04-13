# Decompress Data Using G3ZLIB

## Overview

Decompresses a ZLIB-compressed byte array back to its original binary representation.

## Syntax

```asp
Dim originalData
originalData = obj.Decompress(input)
```

## Parameters and Arguments

- input (Variant, Required): The compressed byte array to decompress.

## Return Values

Returns a byte array (Variant array of bytes) containing the decompressed data. If the decompression fails, it returns Empty and updates the LastError property.

## Remarks

- Method names are case-insensitive.
- Use this method when you expect binary data output. For string output, use DecompressText instead.

## Code Example

```asp
<%
Option Explicit
Dim obj, compressedData, originalData
Set obj = Server.CreateObject("G3ZLIB")

' Assume compressedData is a valid byte array obtained from Compress()
' originalData = obj.Decompress(compressedData)

Set obj = Nothing
%>
```



