# NewContext Method

## Overview
Initializes a new drawing context with specified dimensions in the G3Pix AxonASP G3IMAGE library. This method is also accessible via the "New" alias.

## Syntax
```asp
result = obj.NewContext(Width, Height)
```

## Parameters
- **Width** (Integer): The width of the new context in pixels.
- **Height** (Integer): The height of the new context in pixels.

## Return Values
Returns a Boolean indicating whether the context was successfully created.

## Remarks
- This method maps to both the "New" and "NewContext" method names in the underlying engine.
- Dimensions must be positive integers.
- Calling this method releases any previously active context.

## Code Example
```asp
<%
Dim img
Set img = Server.CreateObject("G3IMAGE")
' Using the alias "New"
If img.New(800, 600) Then
    ' Context is ready
End If
Set img = Nothing
%>
```
