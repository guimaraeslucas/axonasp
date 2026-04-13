# Get the LastError Property

## Overview

The LastError property is exposed by the G3ZLIB library object and returns the last error message recorded during compression or decompression operations.

## Syntax

```asp
Dim errorMessage
errorMessage = obj.LastError
```

## Parameters and Arguments

- Getter: Does not take arguments.
- Setter: This property is read-only.

## Return Values

Returns a String containing the description of the last error that occurred. If no error occurred, it returns an empty string.

## Remarks

- Property names are case-insensitive.
- This property is read-only and rejects assignments.

## Code Example

```asp
<%
Option Explicit
Dim obj, errorMessage
Set obj = Server.CreateObject("G3ZLIB")

' Attempt to decompress invalid data
Call obj.DecompressText("invalid data")

errorMessage = obj.LastError
Response.Write "Last Error: " & errorMessage

Set obj = Nothing
%>
```