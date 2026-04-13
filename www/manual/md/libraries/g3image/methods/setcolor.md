# SetColor Method

## Overview
Sets the active color for drawing and filling operations using an RGB or RGBA string in the G3Pix AxonASP G3IMAGE library.

## Syntax
```asp
obj.SetColor(rgb_string)
```

## Parameters
- **rgb_string** (String): A comma-separated string of color values (e.g., "255,0,0" for red or "0,0,255,128" for semi-transparent blue).

## Return Values
Returns Empty upon completion.

## Remarks
- Color values should be between 0 and 255.
- If four values are provided, the last value is the alpha (transparency) channel.

## Code Example
```asp
<%
Dim img
Set img = Server.CreateObject("G3IMAGE")
If img.NewContext(100, 100) Then
    ' Set to red
    img.SetColor "255,0,0"
    img.DrawCircle 50, 50, 40
    img.Fill()
End If
Set img = Nothing
%>
```
