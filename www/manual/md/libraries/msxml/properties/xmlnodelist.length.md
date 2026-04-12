# XMLNodeList.Length Property

## Overview
Reads or writes the Length property on the MSXML2 XMLNodeList compatibility object.

## Syntax
```asp
Dim obj, value
Set obj = Server.CreateObject("MSXML2.DOMDocument")
value = obj.Length
```

## Access
Read Only

## Return Values
Returns a Variant-compatible value.

## Remarks
- Number of nodes in list.
- Property names are case-insensitive.

## Code Example
```asp
<%
Dim obj
Set obj = Server.CreateObject("MSXML2.DOMDocument")
On Error Resume Next
Response.Write CStr(obj.Length)
On Error GoTo 0
Set obj = Nothing
%>
```