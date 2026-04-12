# DOMDocument.PreserveWhiteSpace Property

## Overview
Reads or writes the PreserveWhiteSpace property on the MSXML2 DOMDocument compatibility object.

## Syntax
```asp
Dim obj, value
Set obj = Server.CreateObject("MSXML2.DOMDocument")
value = obj.PreserveWhiteSpace
```

## Access
Read/Write

## Return Values
Returns a Variant-compatible value.

## Remarks
- Controls whitespace preservation behavior.
- Property names are case-insensitive.

## Code Example
```asp
<%
Dim obj
Set obj = Server.CreateObject("MSXML2.DOMDocument")
On Error Resume Next
Response.Write CStr(obj.PreserveWhiteSpace)
On Error GoTo 0
Set obj = Nothing
%>
```