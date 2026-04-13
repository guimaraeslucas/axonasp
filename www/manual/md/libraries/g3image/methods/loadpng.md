# LoadPNG Method

## Overview
Loads a PNG image file from the specified path into the G3Pix AxonASP G3IMAGE library.

## Syntax
```asp
result = obj.LoadPNG(path)
```

## Parameters
- **path** (String): The file path to the PNG image.

## Return Values
Returns a Boolean indicating whether the image was successfully loaded.

## Remarks
- This method specifically targets PNG files. For general image loading, use the LoadImage method.
- The path is resolved relative to the web root if possible.

## Code Example
```asp
<%
Dim img
Set img = Server.CreateObject("G3IMAGE")
If img.LoadPNG("images/logo.png") Then
    ' PNG loaded successfully
End If
Set img = Nothing
%>
```
