# DOMDocument.XML Property

## Overview
Reads or writes the XML property on the MSXML2 DOMDocument compatibility object.

## Syntax
```asp
Dim obj, value
Set obj = Server.CreateObject("MSXML2.DOMDocument")
value = obj.XML
```

## Access
Read Only

## Return Values
Returns a Variant-compatible value.

## Remarks
- Serialized XML text.
- Property names are case-insensitive.

## Code Example
```asp
<%
Dim obj
Set obj = Server.CreateObject("MSXML2.DOMDocument")
On Error Resume Next
Response.Write CStr(obj.XML)
On Error GoTo 0
Set obj = Nothing
%>
```