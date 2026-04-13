# axintegersizebytes

## Overview
Retrieves the size in bytes of an integer on the current G3Pix AxonASP platform.

## Syntax
```asp
result = obj.axintegersizebytes()
```

## Parameters and Arguments
None.

## Return Values
Returns an Integer representing the number of bytes used to store an integer (typically 4 or 8).

## Remarks
The result depends on whether the process is running on a 32-bit or 64-bit architecture.

## Code Example
```asp
<%
Option Explicit
Dim obj, intSize
Set obj = Server.CreateObject("G3AXON.FUNCTIONS")

intSize = obj.axintegersizebytes()
Response.Write "Integer size on this platform: " & intSize & " bytes"

Set obj = Nothing
%>
```
