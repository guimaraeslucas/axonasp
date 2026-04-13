# Clear Method

## Overview
Clears the current image context using the current fill color. This operation resets the canvas within the active context of the G3Pix AxonASP G3IMAGE library.

## Syntax
```asp
obj.Clear()
```

## Return Values
Returns Empty upon completion.

## Remarks
- Ensure a context has been initialized using NewContext before calling this method.
- The method uses the color previously set by SetColor or SetHexColor.

## Code Example
```asp
<%
Dim img
Set img = Server.CreateObject("G3IMAGE")
If img.NewContext(400, 300) Then
    img.SetHexColor("#FFFFFF")
    img.Clear()
End If
Set img = Nothing
%>
```
