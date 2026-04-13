# SetLevel Method

## Overview

Sets the default compression level for the G3Pix AxonASP G3ZSTD object. This level is used for all subsequent compression operations that do not explicitly specify a level.

## Syntax

```asp
result = obj.SetLevel(level)
```

## Parameters and Arguments

- **level**: An Integer specifying the compression level. The range is -5 (fastest, lower ratio) to 22 (slowest, highest ratio).

## Return Values

Returns a Boolean value:
- **True**: The compression level was successfully updated.
- **False**: An invalid level was provided (outside the -5 to 22 range).

## Remarks

- This method is an alias for `SetCompressionLevel`.
- Setting a new level will re-initialize the internal encoder resource.
- The default compression level is 3.
- If an invalid level is provided, the current level is maintained and a runtime error is raised.

## Code Example

```asp
<%
Option Explicit
Dim objZstd
Set objZstd = Server.CreateObject("G3ZSTD")

' Change default compression level to 12
If objZstd.SetLevel(12) Then
    Response.Write "Default compression level set to " & objZstd.Level
End If

' Use the new default level for compression
Dim compressed
compressed = objZstd.Compress("Test Data")

Set objZstd = Nothing
%>
```
