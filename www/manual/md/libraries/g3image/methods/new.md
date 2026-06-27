# New / NewContext Method

## Overview
Initializes a new drawing context with specified dimensions in the G3Pix AxonASP G3IMAGE library. This method is also accessible via the "New" alias. When using the `Persits.Jpeg` compatibility layer, it supports a third optional parameter to fill the canvas with a background color.

## Syntax
```asp
result = obj.New(Width, Height [, ColorHex])
```

## Parameters
- **Width** (Integer): The width of the new context in pixels.
- **Height** (Integer): The height of the new context in pixels.
- **ColorHex** (String/Integer, Optional): **(Persits.Jpeg Compatibility Only)** The background color for the blank canvas. Can be a hex string (e.g. `"#FFFFFF"`) or a BGR integer value (e.g. `&HFFFFFF`). Defaults to white.

## Return Values
Returns a Boolean indicating whether the context was successfully created.

## Remarks
- This method maps to both the "New" and "NewContext" method names in the underlying engine.
- Dimensions must be positive integers.
- Calling this method releases any previously active context.

## Code Example
```asp
<%
Dim jpeg
Set jpeg = Server.CreateObject("Persits.Jpeg")
' Create a 400x300 canvas with a blue background
If jpeg.New(400, 300, "#0000FF") Then
    ' Context is ready with blue background
End If
jpeg.Close
Set jpeg = Nothing
%>
```
