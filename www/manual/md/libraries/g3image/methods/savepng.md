# SavePNG Method

## Overview
Saves the current context as a PNG file to the specified path in the G3Pix AxonASP G3IMAGE library.

## Syntax
```asp
result = obj.SavePNG(path)
```

## Parameters
- **path** (String): The file path where the PNG will be saved.

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
    img.DrawRectangle 0, 0, 200, 200
    img.Fill()
    img.SavePNG "output.png"
End If
Set img = Nothing
%>
```
