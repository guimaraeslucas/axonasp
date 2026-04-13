# LoadJPG Method

## Overview
Loads a JPEG image file from the specified path into the G3Pix AxonASP G3IMAGE library.

## Syntax
```asp
result = obj.LoadJPG(path)
```

## Parameters
- **path** (String): The file path to the JPEG image.

## Return Values
Returns a Boolean indicating whether the image was successfully loaded.

## Remarks
- This method specifically targets JPEG files. For general image loading, use the LoadImage method.
- The path is resolved relative to the web root if possible.

## Code Example
```asp
<%
Dim img
Set img = Server.CreateObject("G3IMAGE")
If img.LoadJPG("images/photo.jpg") Then
    ' JPEG loaded successfully
End If
Set img = Nothing
%>
```
