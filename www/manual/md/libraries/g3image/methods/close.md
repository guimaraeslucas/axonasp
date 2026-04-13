# Close Method

## Overview
Releases the resources associated with the G3IMAGE object and clears the current context. This method is part of the G3Pix AxonASP G3IMAGE library.

## Syntax
```asp
result = obj.Close()
```

## Return Values
Returns a Boolean indicating whether the resources were successfully released.

## Remarks
- It is good practice to call Close when finished with an image to free memory.
- This method also resets any internal error state.

## Code Example
```asp
<%
Dim img
Set img = Server.CreateObject("G3IMAGE")
If img.NewContext(100, 100) Then
    ' Perform operations
    img.Close()
End If
Set img = Nothing
%>
```
