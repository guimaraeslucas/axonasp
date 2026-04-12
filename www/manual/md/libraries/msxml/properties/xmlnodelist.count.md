# XMLNodeList.Count Property

## Overview
Reads or writes the Count property on the MSXML2 XMLNodeList compatibility object.

## Syntax
```asp
Dim obj, value
Set obj = Server.CreateObject("MSXML2.DOMDocument")
value = obj.Count
```

## Access
Read Only

## Return Values
Returns a Variant-compatible value.

## Remarks
- Alias for list length.
- Property names are case-insensitive.

## Code Example
```asp
<%
Dim obj
Set obj = Server.CreateObject("MSXML2.DOMDocument")
On Error Resume Next
Response.Write CStr(obj.Count)
On Error GoTo 0
Set obj = Nothing
%>
```