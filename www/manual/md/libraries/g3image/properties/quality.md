# Quality Property

## Overview
Gets or sets the JPEG compression quality for save and rendering operations. This property is part of the `Persits.Jpeg` compatibility layer of the G3Pix AxonASP G3IMAGE library.

## Syntax
```asp
' Get Quality
q = obj.Quality

' Set Quality
obj.Quality = newQuality
```

## Return Values
Returns an Integer in the range `0..100`.

## Remarks
- This property is an alias for the native `JPGQuality` property. It controls the compression factor when saving JPEG files or invoking `SendBinary`.

## Code Example
```asp
<%
Dim jpeg
Set jpeg = Server.CreateObject("Persits.Jpeg")
jpeg.Open Server.MapPath("photo.jpg")
jpeg.Quality = 80
jpeg.Save Server.MapPath("compressed.jpg")
jpeg.Close
Set jpeg = Nothing
%>
```
