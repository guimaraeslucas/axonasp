# DrawRectangle Method

## Overview
Adds a rectangle to the current path in the G3Pix AxonASP G3IMAGE library.

## Syntax
```asp
obj.DrawRectangle(x, y, w, h)
```

## Parameters
- **x** (Double): The x-coordinate of the top-left corner.
- **y** (Double): The y-coordinate of the top-left corner.
- **w** (Double): The width of the rectangle.
- **h** (Double): The height of the rectangle.

## Return Values
Returns Empty upon completion.

## Remarks
- This method adds a rectangle to the current path. Call Stroke or Fill to render it.

## Code Example
```asp
<%
Dim img
Set img = Server.CreateObject("G3IMAGE")
If img.NewContext(400, 300) Then
    img.SetHexColor("#0000FF")
    img.DrawRectangle 50, 50, 300, 200
    img.Stroke()
End If
Set img = Nothing
%>
```
