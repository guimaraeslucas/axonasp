# LoadFontFace Method

## Overview
Loads a TrueType font for use in text rendering operations in the G3Pix AxonASP G3IMAGE library.

## Syntax
```asp
result = obj.LoadFontFace(path, points)
```

## Parameters
- **path** (String): The file path to the .ttf font file.
- **points** (Double): The font size in points.

## Return Values
Returns a Boolean indicating whether the font was successfully loaded.

## Remarks
- If a context is active, this method also sets the font for that context.
- Always check the return value or the LastError property if loading fails.

## Code Example
```asp
<%
Dim img
Set img = Server.CreateObject("G3IMAGE")
If img.LoadFontFace("C:\Windows\Fonts\verdanab.ttf", 12) Then
    ' Font loaded successfully
End If
Set img = Nothing
%>
```
