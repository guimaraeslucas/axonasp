# Interpolation Property

## Overview
Gets or sets the resampling algorithm used when resizing images. This property is part of the `Persits.Jpeg` compatibility layer of the G3Pix AxonASP G3IMAGE library.

## Syntax
```asp
' Get Interpolation
alg = obj.Interpolation

' Set Interpolation
obj.Interpolation = algorithmNumber
```

## Return Values
Returns an Integer representing the algorithm.

## Remarks
Supported values:
- `1` (Nearest Neighbor): Fastest speed, lowest quality, blocks/aliasing visible.
- `2` (Bilinear, Default): Moderate speed, good quality for downscaling.
- `3` (Bicubic / Catmull-Rom): Slowest speed, highest quality, sharpest details.

## Code Example
```asp
<%
Dim jpeg
Set jpeg = Server.CreateObject("Persits.Jpeg")
jpeg.Open Server.MapPath("photo.jpg")
' Set Bicubic resampling
jpeg.Interpolation = 3
jpeg.Width = 200
jpeg.Height = 150
jpeg.Save Server.MapPath("resized.jpg")
jpeg.Close
Set jpeg = Nothing
%>
```
