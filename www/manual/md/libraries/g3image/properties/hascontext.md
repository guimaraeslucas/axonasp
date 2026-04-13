# HasContext Property

## Overview
Indicates whether a drawing context has been initialized for the G3Pix AxonASP G3IMAGE library.

## Syntax
```asp
exists = obj.HasContext
```

## Return Values
Returns a Boolean value. True if a context is active; otherwise, False.

## Remarks
- A context is typically created using the NewContext method.
- This property is read-only.

## Code Example
```asp
<%
Dim img
Set img = Server.CreateObject("G3IMAGE")
If Not img.HasContext Then
    img.NewContext 400, 300
End If
Set img = Nothing
%>
```
