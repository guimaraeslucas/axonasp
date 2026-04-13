# Fill Method

## Overview
Fills the current path with the active color in the G3Pix AxonASP G3IMAGE library.

## Syntax
```asp
obj.Fill()
```

## Return Values
Returns Empty upon completion.

## Remarks
- Use SetColor or SetHexColor to define the fill color before calling this method.
- The path is cleared after the fill operation. Use FillPreserve if you need to keep the path for further operations.

## Code Example
```asp
<%
Dim img
Set img = Server.CreateObject("G3IMAGE")
If img.NewContext(200, 200) Then
    img.SetHexColor("#FFCC00")
    img.DrawCircle 100, 100, 80
    img.Fill()
End If
Set img = Nothing
%>
```
