# DrawString Method

## Overview
Draws a string of text at the specified coordinates in the G3Pix AxonASP G3IMAGE library.

## Syntax
```asp
obj.DrawString(s, x, y)
```

## Parameters
- **s** (String): The text to be drawn.
- **x** (Double): The x-coordinate for the text position.
- **y** (Double): The y-coordinate for the text position.

## Return Values
Returns Empty upon completion.

## Remarks
- Use LoadFontFace to set the font and size before drawing strings.
- The coordinates typically refer to the baseline of the text.

## Code Example
```asp
<%
Dim img
Set img = Server.CreateObject("G3IMAGE")
If img.NewContext(400, 200) Then
    img.LoadFontFace "C:\Windows\Fonts\arial.ttf", 24
    img.SetHexColor("#000000")
    img.DrawString "Hello AxonASP", 50, 100
End If
Set img = Nothing
%>
```
