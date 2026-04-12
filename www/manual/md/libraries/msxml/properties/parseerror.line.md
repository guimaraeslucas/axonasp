# ParseError.Line Property

## Overview
Reads or writes the Line property on the MSXML2 ParseError compatibility object.

## Syntax
```asp
Dim obj, value
Set obj = Server.CreateObject("MSXML2.DOMDocument")
value = obj.Line
```

## Access
Read Only

## Return Values
Returns a Variant-compatible value.

## Remarks
- Error line number.
- Property names are case-insensitive.

## Code Example
```asp
<%
Dim obj
Set obj = Server.CreateObject("MSXML2.DOMDocument")
On Error Resume Next
Response.Write CStr(obj.Line)
On Error GoTo 0
Set obj = Nothing
%>
```