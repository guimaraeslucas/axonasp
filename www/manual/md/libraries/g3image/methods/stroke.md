# Stroke Method

## Overview
Outlines the current path using the active color and line width in the G3Pix AxonASP G3IMAGE library.

## Syntax
```asp
obj.Stroke()
```

## Return Values
Returns Empty upon completion.

## Remarks
- Use SetColor or SetHexColor and SetLineWidth to define the appearance before calling this method.
- The path is cleared after the stroke operation. Use StrokePreserve if you need to keep the path.

## Code Example
```asp
<%
Dim img
Set img = Server.CreateObject("G3IMAGE")
If img.NewContext(200, 200) Then
    img.SetHexColor("#000000")
    img.SetLineWidth 2
    img.DrawCircle 100, 100, 80
    img.Stroke()
End If
Set img = Nothing
%>
```
