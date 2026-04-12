# ParseError.Reason Property

## Overview
Reads or writes the Reason property on the MSXML2 ParseError compatibility object.

## Syntax
```asp
Dim obj, value
Set obj = Server.CreateObject("MSXML2.DOMDocument")
value = obj.Reason
```

## Access
Read Only

## Return Values
Returns a Variant-compatible value.

## Remarks
- Parser error message.
- Property names are case-insensitive.

## Code Example
```asp
<%
Dim obj
Set obj = Server.CreateObject("MSXML2.DOMDocument")
On Error Resume Next
Response.Write CStr(obj.Reason)
On Error GoTo 0
Set obj = Nothing
%>
```