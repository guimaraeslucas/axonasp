# axexecutablepath

## Overview
Retrieves the absolute path of the executable that started the G3Pix AxonASP process.

## Syntax
```asp
result = obj.axexecutablepath()
```

## Parameters and Arguments
None.

## Return Values
Returns a String containing the full path to the current executable.

## Remarks
If an error occurs while determining the path, an empty string is returned.

## Code Example
```asp
<%
Option Explicit
Dim obj, execPath
Set obj = Server.CreateObject("G3AXON.FUNCTIONS")

execPath = obj.axexecutablepath()
Response.Write "Current executable path: " & execPath

Set obj = Nothing
%>
```
