# Width Property

## Overview
Gets or sets the width of the active image context.

## Syntax
```asp
' Get width
w = obj.Width

' Set width (resizes the image)
obj.Width = newWidth
```

## Return Values
Returns an Integer representing the width in pixels. Returns 0 if no context is active.

## Remarks
- When using the `Persits.Jpeg` compatibility layer, this property is read/write. Setting it will resample/resize the image to the new width using the algorithm specified by the `Interpolation` property.
- For native G3IMAGE instances, this property remains read-only.

## Code Example
```asp
<%
Dim jpeg
Set jpeg = Server.CreateObject("Persits.Jpeg")
jpeg.Open Server.MapPath("photo.jpg")
jpeg.Width = 200 ' Resizes the image to 200px width
jpeg.Save Server.MapPath("resized.jpg")
jpeg.Close
Set jpeg = Nothing
%>
```
