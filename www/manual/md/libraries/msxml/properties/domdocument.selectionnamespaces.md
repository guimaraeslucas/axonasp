# DOMDocument.SelectionNamespaces Property

## Overview
Reads or writes the SelectionNamespaces property on the MSXML2 DOMDocument compatibility object.

## Syntax
```asp
Dim obj, value
Set obj = Server.CreateObject("MSXML2.DOMDocument")
value = obj.SelectionNamespaces
```

## Access
Read/Write

## Return Values
Returns a Variant-compatible value.

## Remarks
- Namespace mapping string for selection APIs.
- Property names are case-insensitive.

## Code Example
```asp
<%
Dim obj
Set obj = Server.CreateObject("MSXML2.DOMDocument")
On Error Resume Next
Response.Write CStr(obj.SelectionNamespaces)
On Error GoTo 0
Set obj = Nothing
%>
```