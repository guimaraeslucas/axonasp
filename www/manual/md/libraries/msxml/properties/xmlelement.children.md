# XMLElement.Children Property

## Overview
Reads or writes the Children property on the MSXML2 XMLElement compatibility object.

## Syntax
```asp
Dim obj, value
Set obj = Server.CreateObject("MSXML2.DOMDocument")
value = obj.Children
```

## Access
Read Only

## Return Values
Returns a Variant-compatible value.

## Remarks
- Children collection alias.
- Property names are case-insensitive.

## Code Example
```asp
<%
Dim obj
Set obj = Server.CreateObject("MSXML2.DOMDocument")
On Error Resume Next
Response.Write CStr(obj.Children)
On Error GoTo 0
Set obj = Nothing
%>
```