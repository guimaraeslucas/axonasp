# LastBytes Property

## Overview
Returns the binary data of the last rendered or loaded image in the G3Pix AxonASP G3IMAGE library.

## Syntax
```asp
data = obj.LastBytes
```

## Return Values
Returns a VBArray containing binary bytes.

## Remarks
- This property is populated after calling methods like RenderViaTemp.
- It is read-only.

## Code Example
```asp
<%
Dim img, bytes
Set img = Server.CreateObject("G3IMAGE")
img.NewContext 100, 100
img.RenderViaTemp "png", 0
bytes = img.LastBytes
' bytes now contains the binary PNG data
Set img = Nothing
%>
```
