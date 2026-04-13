# Decompress Text Using G3ZLIB

## Overview

Decompresses a ZLIB-compressed byte array back to its original text representation.

## Syntax

```asp
Dim textData
textData = obj.DecompressText(input)
```

## Parameters and Arguments

- input (Variant, Required): The compressed byte array to decompress.

## Return Values

Returns a String containing the decompressed text. If the decompression fails, it returns an empty string and updates the LastError property.

## Remarks

- Method names are case-insensitive.
- This method assumes the original data was a string and automatically handles the text conversion.

## Code Example

```asp
<%
Option Explicit
Dim obj, textData, compressedData
Set obj = Server.CreateObject("G3ZLIB")

compressedData = obj.Compress("Sample text")
textData = obj.DecompressText(compressedData)

Response.Write textData

Set obj = Nothing
%>
```



