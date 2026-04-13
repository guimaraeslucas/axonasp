# DrawCircle Method

## Overview
Adds a circle to the current path in the G3Pix AxonASP G3IMAGE library.

## Syntax
```asp
obj.DrawCircle(x, y, r)
```

## Parameters
- **x** (Double): The x-coordinate of the center of the circle.
- **y** (Double): The y-coordinate of the center of the circle.
- **r** (Double): The radius of the circle.

## Return Values
Returns Empty upon completion.

## Remarks
- This method adds the circle to the current path. You must call Stroke or Fill to actually render the circle on the canvas.

## Code Example
```asp
<%
Dim img
Set img = Server.CreateObject("G3IMAGE")
If img.NewContext(200, 200) Then
    img.SetHexColor("#FF0000")
    img.DrawCircle 100, 100, 50
    img.Stroke()
End If
Set img = Nothing
%>
```
