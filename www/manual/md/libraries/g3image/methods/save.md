# Save Method

## Overview
Saves the modified image to disk. This method is part of the `Persits.Jpeg` compatibility layer of the G3Pix AxonASP G3IMAGE library.

## Syntax
```asp
obj.Save(Path)
```

## Parameters
- `Path`: A string containing the virtual or absolute file path where the image will be saved.

## Return Values
Returns `True` on success. Raises an error on failure.

## Remarks
- If the file extension is `.png`, it will be saved as a PNG; otherwise, it defaults to saving as JPEG using the quality specified by the `Quality` property.

## Code Example
```asp
<%
Dim jpeg
Set jpeg = Server.CreateObject("Persits.Jpeg")
jpeg.Open Server.MapPath("photo.jpg")
jpeg.Width = 100
jpeg.Height = 100
jpeg.Save Server.MapPath("thumbnail.jpg")
jpeg.Close
Set jpeg = Nothing
%>
```
