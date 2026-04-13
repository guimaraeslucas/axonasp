# DrawEllipse Method

## Overview
Adds an ellipse to the current path in the G3Pix AxonASP G3IMAGE library.

## Syntax
```asp
obj.DrawEllipse(x, y, rx, ry)
```

## Parameters
- **x** (Double): The x-coordinate of the center of the ellipse.
- **y** (Double): The y-coordinate of the center of the ellipse.
- **rx** (Double): The horizontal radius of the ellipse.
- **ry** (Double): The vertical radius of the ellipse.

## Return Values
Returns Empty upon completion.

## Remarks
- This method adds the ellipse to the current path. You must call Stroke or Fill to actually render it on the canvas.

## Code Example
```asp
<%
Dim img
Set img = Server.CreateObject("G3IMAGE")
If img.NewContext(200, 200) Then
    img.SetHexColor("#00FF00")
    img.DrawEllipse 100, 100, 80, 40
    img.Fill()
End If
Set img = Nothing
%>
```
