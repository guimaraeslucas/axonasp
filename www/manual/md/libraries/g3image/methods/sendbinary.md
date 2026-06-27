# SendBinary Method

## Overview
Returns the image as a binary byte array. This method is part of the `Persits.Jpeg` compatibility layer of the G3Pix AxonASP G3IMAGE library.

## Syntax
```asp
bytes = obj.SendBinary()
```

## Return Values
Returns a one-dimensional array of bytes (VTArray/VTUI1) representing the image.

## Remarks
- The returned byte array can be passed directly to `Response.BinaryWrite` to stream the image without saving it to disk.

## Code Example
```asp
<%
Dim jpeg, bytes
Set jpeg = Server.CreateObject("Persits.Jpeg")
jpeg.Open Server.MapPath("photo.jpg")
jpeg.Width = 200
jpeg.Height = 150
bytes = jpeg.SendBinary()
Response.ContentType = "image/jpeg"
Response.BinaryWrite bytes
jpeg.Close
Set jpeg = Nothing
%>
```
