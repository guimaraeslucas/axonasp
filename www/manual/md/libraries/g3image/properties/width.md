# Width Property

## Overview
Returns the width of the current drawing context in the G3Pix AxonASP G3IMAGE library.

## Syntax
```asp
w = obj.Width
```

## Return Values
Returns an Integer representing the width in pixels. Returns 0 if no context is active.

## Remarks
- This property is read-only. To set the width, create a new context using NewContext.

## Code Example
```asp
<%
Dim img
Set img = Server.CreateObject("G3IMAGE")
img.NewContext 400, 300
Response.Write "Context width: " & img.Width
Set img = Nothing
%>
```
