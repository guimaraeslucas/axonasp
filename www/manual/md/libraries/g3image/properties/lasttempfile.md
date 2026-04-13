# LastTempFile Property

## Overview
Returns the file path of the temporary file used during the last rendering operation in the G3Pix AxonASP G3IMAGE library.

## Syntax
```asp
path = obj.LastTempFile
```

## Return Values
Returns a String containing the full file path.

## Remarks
- The file is typically deleted automatically after the operation completes.
- This property is primarily used for debugging purposes.
- It is read-only.

## Code Example
```asp
<%
Dim img
Set img = Server.CreateObject("G3IMAGE")
img.NewContext 100, 100
img.RenderViaTemp "png", 0
Response.Write "Temp file used: " & img.LastTempFile
Set img = Nothing
%>
```
