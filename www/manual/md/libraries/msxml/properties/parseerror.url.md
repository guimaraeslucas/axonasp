# ParseError.URL Property

## Overview
Reads or writes the URL property on the MSXML2 ParseError compatibility object.

## Syntax
```asp
Dim obj, value
Set obj = Server.CreateObject("MSXML2.DOMDocument")
value = obj.URL
```

## Access
Read Only

## Return Values
Returns a Variant-compatible value.

## Remarks
- Source URL/path for parser error.
- Property names are case-insensitive.

## Code Example
```asp
<%
Dim obj
Set obj = Server.CreateObject("MSXML2.DOMDocument")
On Error Resume Next
Response.Write CStr(obj.URL)
On Error GoTo 0
Set obj = Nothing
%>
```