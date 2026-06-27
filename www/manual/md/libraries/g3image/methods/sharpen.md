# Sharpen Method

## Overview
Applies a sharpening filter to the image. This method is part of the `Persits.Jpeg` compatibility layer of the G3Pix AxonASP G3IMAGE library.

## Syntax
```asp
obj.Sharpen(Radius, Amount)
```

## Parameters
- `Radius`: Double value representing the radius of the sharpening kernel.
- `Amount`: Double value representing the amount of sharpening (often expressed as percentage, e.g. 120 or 1.2).

## Return Values
Returns Empty.

## Code Example
```asp
<%
Dim jpeg
Set jpeg = Server.CreateObject("Persits.Jpeg")
jpeg.Open Server.MapPath("photo.jpg")
jpeg.Sharpen 1.5, 120
jpeg.Save Server.MapPath("sharp.jpg")
jpeg.Close
Set jpeg = Nothing
%>
```
