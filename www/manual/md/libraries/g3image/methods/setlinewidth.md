# SetLineWidth Method

## Overview
Sets the width of the line for subsequent stroke operations in the G3Pix AxonASP G3IMAGE library.

## Syntax
```asp
obj.SetLineWidth(width)
```

## Parameters
- **width** (Double): The thickness of the line in pixels.

## Return Values
Returns Empty upon completion.

## Remarks
- This setting affects all future Stroke calls until it is changed again.

## Code Example
```asp
<%
Dim img
Set img = Server.CreateObject("G3IMAGE")
If img.NewContext(100, 100) Then
    img.SetLineWidth 5.5
    img.DrawLine 0, 0, 100, 100
    img.Stroke()
End If
Set img = Nothing
%>
```
