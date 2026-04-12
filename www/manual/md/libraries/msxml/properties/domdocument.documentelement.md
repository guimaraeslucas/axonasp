# DOMDocument.DocumentElement Property

## Overview
Reads or writes the DocumentElement property on the MSXML2 DOMDocument compatibility object.

## Syntax
```asp
Dim obj, value
Set obj = Server.CreateObject("MSXML2.DOMDocument")
value = obj.DocumentElement
```

## Access
Read Only

## Return Values
Returns a Variant-compatible value.

## Remarks
- Root element node.
- Property names are case-insensitive.

## Code Example
```asp
<%
Dim obj
Set obj = Server.CreateObject("MSXML2.DOMDocument")
On Error Resume Next
Response.Write CStr(obj.DocumentElement)
On Error GoTo 0
Set obj = Nothing
%>
```