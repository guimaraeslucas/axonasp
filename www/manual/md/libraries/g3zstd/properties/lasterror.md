# LastError Property

## Overview

Returns a String describing the last error that occurred during a Zstandard (zstd) operation in the G3Pix AxonASP G3ZSTD object.

## Syntax

```asp
errorString = obj.LastError
```

## Parameters and Arguments

This property is read-only and does not accept any parameters.

## Return Values

Returns a String containing the error message associated with the last failed operation. If no error has occurred, it returns an empty string.

## Remarks

- This property is useful for diagnostic purposes when a method (like `CompressFile`) returns `False`.
- The error message is reset when `Clear` is called or when a new operation succeeds.
- Runtime exceptions also contain these messages.

## Code Example

```asp
<%
Option Explicit
Dim objZstd
Set objZstd = Server.CreateObject("G3ZSTD")

' Perform an operation that might fail
If Not objZstd.CompressFile("missing_file.txt", "target.zst") Then
    Response.Write "Operation failed with error: " & objZstd.LastError
End If

Set objZstd = Nothing
%>
```
