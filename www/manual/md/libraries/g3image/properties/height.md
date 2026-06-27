# Height Property

## Overview
Gets or sets the height of the active image context.

## Syntax
```asp
' Get height
h = obj.Height

' Set height (resizes the image)
obj.Height = newHeight
```

## Return Values
Returns an Integer representing the height in pixels. Returns 0 if no context is active.

## Remarks
- When using the `Persits.Jpeg` compatibility layer, this property is read/write. Setting it will resample/resize the image to the new height using the algorithm specified by the `Interpolation` property.
- For native G3IMAGE instances, this property remains read-only.

## Code Example
```asp
<%
Dim jpeg
Set jpeg = Server.CreateObject("Persits.Jpeg")
jpeg.Open Server.MapPath("photo.jpg")
jpeg.Height = 150 ' Resizes the image to 150px height
jpeg.Save Server.MapPath("resized.jpg")
jpeg.Close
Set jpeg = Nothing
%>
```
