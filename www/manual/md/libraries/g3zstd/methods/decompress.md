# Decompress Method

## Overview

Decompresses a Zstandard (zstd) compressed payload (string or byte array) and returns the original data as a byte array.

## Syntax

```asp
result = obj.Decompress(input)
```

## Parameters and Arguments

- **input**: A String or a VBArray of bytes containing the Zstandard compressed data.

## Return Values

Returns a VBArray of bytes containing the decompressed data. If an error occurs, it returns Empty.

## Remarks

- The input must be a valid Zstandard frame.
- This G3Pix AxonASP method handles all data types by normalizing them to a byte stream before decompression.
- If decompression fails (e.g., due to corrupt data or invalid format), a runtime error is raised.

## Code Example

```asp
<%
Option Explicit
Dim objZstd, compressedData, originalData
' Assuming compressedData is obtained from a source or previous operation
Set objZstd = Server.CreateObject("G3ZSTD")

originalData = objZstd.Decompress(compressedData)

If Not IsEmpty(originalData) Then
    Response.Write "Original data size: " & UBound(originalData) + 1 & " bytes."
Else
    Response.Write "Decompression failed."
End If

Set objZstd = Nothing
%>
```
