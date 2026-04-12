# XMLElement.LastChild Property

## Overview
Reads or writes the LastChild property on the MSXML2 XMLElement compatibility object.

## Syntax
```asp
Dim obj, value
Set obj = Server.CreateObject("MSXML2.DOMDocument")
value = obj.LastChild
```

## Access
Read Only

## Return Values
Returns a Variant-compatible value.

## Remarks
- Last child node.
- Property names are case-insensitive.

## Code Example
```asp
<%
Dim obj
Set obj = Server.CreateObject("MSXML2.DOMDocument")
On Error Resume Next
Response.Write CStr(obj.LastChild)
On Error GoTo 0
Set obj = Nothing
%>
```