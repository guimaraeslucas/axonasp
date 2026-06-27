# Crop Method

## Overview
Crops the image to the specified coordinates. This method is part of the `Persits.Jpeg` compatibility layer of the G3Pix AxonASP G3IMAGE library.

## Syntax
```asp
obj.Crop(x0, y0, x1, y1)
```

## Parameters
- `x0`, `y0`: Coordinates of the top-left corner of the cropping rectangle.
- `x1`, `y1`: Coordinates of the bottom-right corner of the cropping rectangle.

## Return Values
Returns `True` on success. Raises an error on failure.

## Code Example
```asp
<%
Dim jpeg
Set jpeg = Server.CreateObject("Persits.Jpeg")
jpeg.Open Server.MapPath("photo.jpg")
jpeg.Crop 10, 10, 100, 100
jpeg.Save Server.MapPath("cropped.jpg")
jpeg.Close
Set jpeg = Nothing
%>
```
