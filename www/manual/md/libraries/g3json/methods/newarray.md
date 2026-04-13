# NewArray Method

## Overview
Creates a new, empty VBScript-compatible array.

## Syntax
```asp
arr = json.NewArray()
```

## Return Values
Returns an **Array** (Variant) containing zero elements.

## Remarks
This method provides a convenient way to initialize an empty array that can be used for building data structures to be stringified via the **Stringify** method.

## Code Example
```asp
<%
Dim json, arr
Set json = Server.CreateObject("G3JSON")
arr = json.NewArray()
' Array is ready for use
Response.Write "Array created, size: " & UBound(arr) + 1
Set json = Nothing
%>
```
