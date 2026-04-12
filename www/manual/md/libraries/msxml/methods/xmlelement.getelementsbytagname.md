# XMLElement.GetElementsByTagName Method

## Overview
Calls the GetElementsByTagName member on the MSXML2 XMLElement compatibility object.

## Syntax
```asp
Dim obj
Set obj = Server.CreateObject("MSXML2.DOMDocument")
Set list = node.GetElementsByTagName("child")
```

## Parameters and Arguments
- Parameters are validated by runtime dispatch for this object.
- Invalid argument count or incompatible values can raise runtime errors.

## Return Values
Returns a Variant-compatible value or native object handle depending on the operation.

## Remarks
- Finds descendant elements by tag name.
- Member names are case-insensitive.
- Use Set for object return values.

## Code Example
```asp
<%
Dim obj
Set obj = Server.CreateObject("MSXML2.DOMDocument")
On Error Resume Next
obj.GetElementsByTagName
If Err.Number <> 0 Then
    Response.Write "Error: " & Err.Description
    Err.Clear
End If
On Error GoTo 0
Set obj = Nothing
%>
```