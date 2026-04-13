# LastError Property

## Overview
Returns the error message for the last operation performed in the G3Pix AxonASP G3IMAGE library.

## Syntax
```asp
err = obj.LastError
```

## Return Values
Returns a String containing the last error message. Returns an empty string if no error occurred.

## Remarks
- Check this property after any operation that returns False or Empty to understand why it failed.
- This property is read-only.

## Code Example
```asp
<%
Dim img
Set img = Server.CreateObject("G3IMAGE")
If Not img.LoadImage("missing.png") Then
    Response.Write "Error: " & img.LastError
End If
Set img = Nothing
%>
```
