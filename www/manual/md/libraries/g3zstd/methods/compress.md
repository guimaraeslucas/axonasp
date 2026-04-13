# Compress Method

## Overview

Compresses a string or a byte array using the Zstandard (zstd) algorithm. This G3Pix AxonASP method provides high-performance compression with adjustable levels.

## Syntax

```asp
result = obj.Compress(input [, level])
```

## Parameters and Arguments

- **input**: A String or a VBArray of bytes containing the data to be compressed.
- **level** (Optional): An Integer specifying the compression level. The range is -5 to 22. If omitted, the default level (currently 3) is used.

## Return Values

Returns a VBArray of bytes containing the compressed data. If an error occurs, it returns Empty.

## Remarks

- The compression level can range from -5 (fastest) to 22 (highest compression ratio).
- Higher compression levels consume more memory and CPU.
- Input validation ensures that the input is either a string or a valid array of bytes (0-255).
- If the compression level is invalid, a runtime error is raised.

## Code Example

```asp
<%
Option Explicit
Dim objZstd, compressedData, sourceString
sourceString = "This is some sample text to compress using G3Pix AxonASP G3ZSTD."

Set objZstd = Server.CreateObject("G3ZSTD")

' Compress using default level
compressedData = objZstd.Compress(sourceString)

If Not IsEmpty(compressedData) Then
    Response.Write "Compressed data size: " & UBound(compressedData) + 1 & " bytes."
Else
    Response.Write "Compression failed: " & objZstd.LastError
End If

Set objZstd = Nothing
%>
```
