# DrawStringAnchored Method

## Overview
Draws a string of text with specified anchor points for alignment in the G3Pix AxonASP G3IMAGE library.

## Syntax
```asp
obj.DrawStringAnchored(s, x, y, ax, ay)
```

## Parameters
- **s** (String): The text to be drawn.
- **x** (Double): The x-coordinate for the anchor position.
- **y** (Double): The y-coordinate for the anchor position.
- **ax** (Double): The horizontal anchor point (0.0 for left, 0.5 for center, 1.0 for right).
- **ay** (Double): The vertical anchor point (0.0 for top, 0.5 for middle, 1.0 for bottom).

## Return Values
Returns Empty upon completion.

## Remarks
- This method allows for precise alignment of text relative to a point.
- Use LoadFontFace to set the font and size before drawing.

## Code Example
```asp
<%
Dim img
Set img = Server.CreateObject("G3IMAGE")
If img.NewContext(400, 200) Then
    img.LoadFontFace "C:\Windows\Fonts\arial.ttf", 24
    img.SetHexColor("#000000")
    ' Draw centered text
    img.DrawStringAnchored "Centered Text", 200, 100, 0.5, 0.5
End If
Set img = Nothing
%>
```
