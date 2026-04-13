# LoadImage Method

## Overview
Loads an image file from the specified path into the G3Pix AxonASP G3IMAGE library.

## Syntax
```asp
result = obj.LoadImage(path)
```

## Parameters
- **path** (String): The file path to the image. Supports PNG, JPEG, and GIF formats.

## Return Values
Returns a Boolean indicating whether the image was successfully loaded.

## Remarks
- After loading an image, you can use DrawImage to render it into a context.
- The path is resolved relative to the web root if possible.

## Code Example
```asp
<%
Dim img
Set img = Server.CreateObject("G3IMAGE")
If img.LoadImage("images/header.jpg") Then
    ' Image loaded
Else
    Response.Write "Error: " & img.LastError
End If
Set img = Nothing
%>
```
