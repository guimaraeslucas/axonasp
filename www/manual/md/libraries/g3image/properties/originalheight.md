# OriginalHeight Property

## Overview
Returns the original height of the opened image. This property is part of the `Persits.Jpeg` compatibility layer of the G3Pix AxonASP G3IMAGE library.

## Syntax
```asp
h = obj.OriginalHeight
```

## Return Values
Returns an Integer representing the original height in pixels.

## Remarks
- This property is read-only. It retains the dimensions of the image as it was loaded/opened, even if the image is subsequently resized via the `Width` or `Height` properties.

## Code Example
```asp
<%
Dim jpeg
Set jpeg = Server.CreateObject("Persits.Jpeg")
jpeg.Open Server.MapPath("photo.jpg")
Response.Write "Original: " & jpeg.OriginalHeight & ", Current: " & jpeg.Height
jpeg.Height = jpeg.OriginalHeight / 2
Response.Write "After half-resize, Current height: " & jpeg.Height
jpeg.Close
Set jpeg = Nothing
%>
```
