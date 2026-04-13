# axeffectiveuserid

## Overview
Retrieves the effective user ID of the current G3Pix AxonASP process.

## Syntax
```asp
result = obj.axeffectiveuserid()
```

## Parameters and Arguments
None.

## Return Values
Returns an Integer representing the effective user ID (euid). On Windows systems, this method always returns -1.

## Remarks
This method is primarily useful on Unix-like systems where it returns the numeric user ID of the effective user.

## Code Example
```asp
<%
Option Explicit
Dim obj, euid
Set obj = Server.CreateObject("G3AXON.FUNCTIONS")

euid = obj.axeffectiveuserid()

If euid = -1 Then
    Response.Write "Effective User ID is not applicable on this system (Windows)."
Else
    Response.Write "Effective User ID: " & euid
End If

Set obj = Nothing
%>
```
