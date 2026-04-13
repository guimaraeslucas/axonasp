# DefaultFormat Property

## Overview
Gets or sets the default image format used for rendering operations in the G3Pix AxonASP G3IMAGE library.

## Syntax
```asp
' Get the current format
format = obj.DefaultFormat

' Set a new format
obj.DefaultFormat = "jpg"
```

## Return Values
Returns a String containing the current default format (e.g., "png" or "jpg").

## Remarks
- The default value is "png".
- Supported values are "png", "jpg", and "jpeg".

## Code Example
```asp
<%
Dim img
Set img = Server.CreateObject("G3IMAGE")
img.DefaultFormat = "jpg"
Response.Write "Default format: " & img.DefaultFormat
Set img = Nothing
%>
```
