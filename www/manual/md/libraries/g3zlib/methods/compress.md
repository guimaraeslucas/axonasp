# Compress Data Using G3ZLIB

## Overview

Compresses text or binary data into a ZLIB compressed format.

## Syntax

```asp
Dim compressedData
compressedData = obj.Compress(input, level)
```

## Parameters and Arguments

- input (Variant, Required): The string or byte array to compress.
- level (Integer, Optional): The compression level from 1 (fastest) to 9 (best compression). If omitted, it uses the default compression level.

## Return Values

Returns a byte array (Variant array of bytes) containing the compressed data. If the compression fails, it returns Empty and updates the LastError property.

## Remarks

- Method names are case-insensitive.
- Always check if the returned value is Empty to verify if the operation succeeded.

## Code Example

```asp
<%
Option Explicit
Dim obj, compressedData
Set obj = Server.CreateObject("G3ZLIB")
compressedData = obj.Compress("This is a sample text to compress.", 9)

If IsEmpty(compressedData) Then
    Response.Write "Compression failed: " & obj.LastError
Else
    Response.Write "Compression succeeded."
End If

Set obj = Nothing
%>
```



