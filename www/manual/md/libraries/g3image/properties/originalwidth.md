# OriginalWidth Property

## Overview
Returns the original width of the opened image. This property is part of the `Persits.Jpeg` compatibility layer of the G3Pix AxonASP G3IMAGE library.

## Syntax
```asp
w = obj.OriginalWidth
```

## Return Values
Returns an Integer representing the original width in pixels.

## Remarks
- This property is read-only. It retains the dimensions of the image as it was loaded/opened, even if the image is subsequently resized via the `Width` or `Height` properties.

## Code Example
```asp
<%
Dim jpeg
Set jpeg = Server.CreateObject("Persits.Jpeg")
jpeg.Open Server.MapPath("photo.jpg")
Response.Write "Original: " & jpeg.OriginalWidth & ", Current: " & jpeg.Width
jpeg.Width = jpeg.OriginalWidth / 2
Response.Write "After half-resize, Current width: " & jpeg.Width
jpeg.Close
Set jpeg = Nothing
%>
```
