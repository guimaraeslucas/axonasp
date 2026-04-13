# Height Property

## Overview
Returns the height of the current drawing context in the G3Pix AxonASP G3IMAGE library.

## Syntax
```asp
h = obj.Height
```

## Return Values
Returns an Integer representing the height in pixels. Returns 0 if no context is active.

## Remarks
- This property is read-only. To set the height, create a new context using NewContext.

## Code Example
```asp
<%
Dim img
Set img = Server.CreateObject("G3IMAGE")
img.NewContext 400, 300
Response.Write "Context height: " & img.Height
Set img = Nothing
%>
```
