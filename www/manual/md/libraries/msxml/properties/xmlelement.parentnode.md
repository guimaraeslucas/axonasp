# XMLElement.ParentNode Property

## Overview
Reads or writes the ParentNode property on the MSXML2 XMLElement compatibility object.

## Syntax
```asp
Dim obj, value
Set obj = Server.CreateObject("MSXML2.DOMDocument")
value = obj.ParentNode
```

## Access
Read Only

## Return Values
Returns a Variant-compatible value.

## Remarks
- Parent node reference.
- Property names are case-insensitive.

## Code Example
```asp
<%
Dim obj
Set obj = Server.CreateObject("MSXML2.DOMDocument")
On Error Resume Next
Response.Write CStr(obj.ParentNode)
On Error GoTo 0
Set obj = Nothing
%>
```