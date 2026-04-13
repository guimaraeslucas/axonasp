# RenderViaTemp Method

## Overview
Renders the current context into a temporary file and returns the binary content in the G3Pix AxonASP G3IMAGE library.

## Syntax
```asp
result = obj.RenderViaTemp(format, quality)
```

## Parameters
- **format** (String): The target format, such as "png" or "jpg".
- **quality** (Integer): The quality level (1-100), primarily for JPEG output.

## Return Values
Returns a VBArray containing binary bytes. This array is compatible with Response.BinaryWrite.

## Remarks
- This method is useful for streaming the generated image directly to the client browser.
- It handles the creation and cleanup of temporary storage automatically.

## Code Example
```asp
<%
Dim img, bytes
Set img = Server.CreateObject("G3IMAGE")
If img.NewContext(100, 100) Then
    img.SetHexColor("#FF0000")
    img.Clear()
    bytes = img.RenderViaTemp("png", 0)
    Response.ContentType = "image/png"
    Response.BinaryWrite bytes
End If
Set img = Nothing
%>
```
