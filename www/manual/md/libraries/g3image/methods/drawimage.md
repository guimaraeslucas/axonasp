# DrawImage Method

## Overview
Draws an image onto the current context in the G3Pix AxonASP G3IMAGE library. Supports both native G3IMAGE rendering and the `Persits.Jpeg` compatible method signature.

## Syntax
```asp
' G3IMAGE syntax
obj.DrawImage(x, y)

' Persits.Jpeg syntax
obj.DrawImage(x, y, JpegObject)
```

## Parameters
- **x** (Integer): The x-coordinate where the top-left corner of the image will be placed.
- **y** (Integer): The y-coordinate where the top-left corner of the image will be placed.
- **JpegObject** (Object, Optional): **(Persits.Jpeg Compatibility Only)** Another `Persits.Jpeg` compatibility object instance whose active image will be drawn onto the current context.

## Return Values
Returns Empty upon completion.

## Code Example
```asp
<%
Dim jpeg1, jpeg2
Set jpeg1 = Server.CreateObject("Persits.Jpeg")
Set jpeg2 = Server.CreateObject("Persits.Jpeg")

jpeg1.Open Server.MapPath("background.jpg")
jpeg2.Open Server.MapPath("logo.png")

' Draw logo onto background at coordinates (50, 50)
jpeg1.DrawImage 50, 50, jpeg2

jpeg1.Save Server.MapPath("composed.jpg")
jpeg1.Close
jpeg2.Close
Set jpeg1 = Nothing
Set jpeg2 = Nothing
%>
```
