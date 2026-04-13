# LastMimeType Property

## Overview
Returns the MIME type of the last image rendered or loaded in the G3Pix AxonASP G3IMAGE library.

## Syntax
```asp
mime = obj.LastMimeType
```

## Return Values
Returns a String containing the MIME type (e.g., "image/png" or "image/jpeg").

## Remarks
- This property can be used to set the `Response.ContentType` before streaming binary content.
- It is read-only.

## Code Example
```asp
<%
Dim img
Set img = Server.CreateObject("G3IMAGE")
img.NewContext 100, 100
img.RenderViaTemp "png", 0
Response.ContentType = img.LastMimeType
Response.BinaryWrite img.LastBytes
Set img = Nothing
%>
```
