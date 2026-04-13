# MeasureString Method

## Overview
Measures the width and height of a string if it were rendered with the current font in the G3Pix AxonASP G3IMAGE library.

## Syntax
```asp
result = obj.MeasureString(s)
```

## Parameters
- **s** (String): The text to measure.

## Return Values
Returns a VBArray containing two Doubles: the width [0] and the height [1] of the string.

## Remarks
- A context must be active and a font must be loaded using LoadFontFace for this method to provide accurate measurements.

## Code Example
```asp
<%
Dim img, size
Set img = Server.CreateObject("G3IMAGE")
If img.NewContext(400, 200) Then
    img.LoadFontFace "C:\Windows\Fonts\arial.ttf", 12
    size = img.MeasureString("Hello")
    Response.Write "Width: " & size(0) & ", Height: " & size(1)
End If
Set img = Nothing
%>
```
