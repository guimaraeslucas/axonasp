# DOMDocument.ParseError Property

## Overview
Reads or writes the ParseError property on the MSXML2 DOMDocument compatibility object.

## Syntax
```asp
Dim obj, value
Set obj = Server.CreateObject("MSXML2.DOMDocument")
value = obj.ParseError
```

## Access
Read Only

## Return Values
Returns a Variant-compatible value.

## Remarks
- ParseError object from latest parse/load operation.
- Property names are case-insensitive.

## Code Example
```asp
<%
Dim obj
Set obj = Server.CreateObject("MSXML2.DOMDocument")
On Error Resume Next
Response.Write CStr(obj.ParseError)
On Error GoTo 0
Set obj = Nothing
%>
```