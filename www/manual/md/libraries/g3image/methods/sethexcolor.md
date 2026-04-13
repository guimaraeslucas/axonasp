# SetHexColor Method

## Overview
Sets the active color for drawing and filling operations using a hexadecimal string in the G3Pix AxonASP G3IMAGE library.

## Syntax
```asp
obj.SetHexColor(hex_string)
```

## Parameters
- **hex_string** (String): A hexadecimal color string (e.g., "#FF0000" or "00FF00"). The "#" prefix is optional.

## Return Values
Returns Empty upon completion.

## Remarks
- Supports 3, 4, 6, and 8 digit hex codes.
- 4 and 8 digit codes include an alpha (transparency) channel.

## Code Example
```asp
<%
Dim img
Set img = Server.CreateObject("G3IMAGE")
If img.NewContext(100, 100) Then
    ' Set to blue
    img.SetHexColor "#0000FF"
    img.DrawRectangle 10, 10, 80, 80
    img.Stroke()
End If
Set img = Nothing
%>
```
