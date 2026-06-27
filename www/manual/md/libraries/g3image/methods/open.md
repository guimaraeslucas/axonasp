# Open Method

## Overview
Loads an image from the local disk and prepares the context. This method is part of the `Persits.Jpeg` compatibility layer of the G3Pix AxonASP G3IMAGE library.

## Syntax
```asp
obj.Open(Path)
```

## Parameters
- `Path`: A string containing the virtual or absolute file path to the image to open.

## Return Values
Returns `True` on success. Raises an error on failure.

## Remarks
- Absolute paths and mapped virtual paths using Server.MapPath are supported.
- Sets the `OriginalWidth` and `OriginalHeight` properties.

## Code Example
```asp
<%
Dim jpeg
Set jpeg = Server.CreateObject("Persits.Jpeg")
jpeg.Open Server.MapPath("photo.jpg")
Response.Write "Opened image: " & jpeg.Width & "x" & jpeg.Height
jpeg.Close
Set jpeg = Nothing
%>
```
