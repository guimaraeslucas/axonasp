# DecompressText Method

## Overview

Decompresses a Zstandard (zstd) compressed payload (string or byte array) and returns the original data as a UTF-8 string.

## Syntax

```asp
resultString = obj.DecompressText(input)
```

## Parameters and Arguments

- **input**: A String or a VBArray of bytes containing the Zstandard compressed data.

## Return Values

Returns a String containing the decompressed UTF-8 text. If an error occurs, it returns an empty string.

## Remarks

- This method is an alias for `DecompressString`.
- It expects the original uncompressed data to be a valid UTF-8 string.
- If decompression fails or the data is not valid UTF-8, it returns an empty string and logs the error to the `LastError` property.

## Code Example

```asp
<%
Option Explicit
Dim objZstd, compressedData, originalText
' Assuming compressedData is obtained from a source or previous operation
Set objZstd = Server.CreateObject("G3ZSTD")

originalText = objZstd.DecompressText(compressedData)

If originalText <> "" Then
    Response.Write "Restored text: " & originalText
Else
    Response.Write "Decompression failed or returned empty text."
End If

Set objZstd = Nothing
%>
```
