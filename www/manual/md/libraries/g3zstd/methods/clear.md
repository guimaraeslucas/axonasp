# Clear Method

## Overview

Clears the internal compression and decompression context of the G3Pix AxonASP G3ZSTD object. This method closes any active encoder or decoder resources and resets the last error message.

## Syntax

```asp
result = obj.Clear()
```

## Parameters and Arguments

This method does not accept any parameters.

## Return Values

Returns a Boolean value:
- **True**: The internal state was successfully cleared.

## Remarks

- Use this method to free up system resources if the object is being reused for multiple unrelated operations.
- This method is also aliased as `Initialize` and `Dispose`.
- If an error occurs during clearing, a runtime error is raised.

## Code Example

```asp
<%
Option Explicit
Dim objZstd
Set objZstd = Server.CreateObject("G3ZSTD")

' Perform operations...

' Clear the internal state
If objZstd.Clear() Then
    Response.Write "Context cleared successfully."
End If

Set objZstd = Nothing
%>
```
