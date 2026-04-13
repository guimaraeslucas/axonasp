# SaveJPG Method

## Overview
Saves the current context as a JPEG file to the specified path in the G3Pix AxonASP G3IMAGE library.

## Syntax
```asp
result = obj.SaveJPG(path, quality)
```

## Parameters
- **path** (String): The file path where the JPEG will be saved.
- **quality** (Integer): The quality level (1-100) for the JPEG compression.

## Return Values
Returns a Boolean indicating whether the file was successfully saved.

## Remarks
- Ensure that the application has write permissions for the target directory.
- The path is resolved relative to the web root if possible.

## Code Example
```asp
<%
Dim img
Set img = Server.CreateObject("G3IMAGE")
If img.NewContext(200, 200) Then
    img.DrawCircle 100, 100, 50
    img.Stroke()
    img.SaveJPG "output.jpg", 85
End If
Set img = Nothing
%>
```
