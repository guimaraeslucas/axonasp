# DrawImage Method

## Overview
Draws the last loaded image onto the current context in the G3Pix AxonASP G3IMAGE library.

## Syntax
```asp
obj.DrawImage(x, y)
```

## Parameters
- **x** (Integer): The x-coordinate where the top-left corner of the image will be placed.
- **y** (Integer): The y-coordinate where the top-left corner of the image will be placed.

## Return Values
Returns Empty upon completion.

## Remarks
- You must load an image using LoadImage, LoadPNG, or LoadJPG before calling this method.
- The image is drawn directly onto the active context.

## Code Example
```asp
<%
Dim img
Set img = Server.CreateObject("G3IMAGE")
If img.LoadImage("assets/logo.png") Then
    If img.NewContext(500, 500) Then
        img.DrawImage 10, 10
    End If
End If
Set img = Nothing
%>
```
