# DrawLine Method

## Overview
Adds a line segment to the current path in the G3Pix AxonASP G3IMAGE library.

## Syntax
```asp
obj.DrawLine(x1, y1, x2, y2)
```

## Parameters
- **x1** (Double): The x-coordinate of the starting point.
- **y1** (Double): The y-coordinate of the starting point.
- **x2** (Double): The x-coordinate of the ending point.
- **y2** (Double): The y-coordinate of the ending point.

## Return Values
Returns Empty upon completion.

## Remarks
- This method defines a line in the current path. Call Stroke to render the line.

## Code Example
```asp
<%
Dim img
Set img = Server.CreateObject("G3IMAGE")
If img.NewContext(200, 200) Then
    img.SetHexColor("#000000")
    img.SetLineWidth 2
    img.DrawLine 10, 10, 190, 190
    img.Stroke()
End If
Set img = Nothing
%>
```
