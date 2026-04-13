# Level Property

## Overview

Gets the current default compression level for the G3Pix AxonASP G3ZSTD object.

## Syntax

```asp
currentLevel = obj.Level
```

## Parameters and Arguments

This property is read-only. To set the level, use the `SetLevel` method.

## Return Values

Returns an Integer representing the current compression level. The range is -5 to 22.

## Remarks

- This property is an alias for `CompressionLevel`.
- The default compression level is 3.
- This level is used for all compression methods unless a level is explicitly provided as an argument.
- Zstandard levels: -5 (fastest) up to 22 (best compression).

## Code Example

```asp
<%
Option Explicit
Dim objZstd
Set objZstd = Server.CreateObject("G3ZSTD")

' Read default level
Response.Write "Default level is " & objZstd.Level

' Change level
objZstd.SetLevel(15)
Response.Write "New level is " & objZstd.Level

Set objZstd = Nothing
%>
```
