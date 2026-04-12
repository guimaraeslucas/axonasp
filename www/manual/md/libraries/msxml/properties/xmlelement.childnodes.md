# XMLElement.ChildNodes Property

## Overview
Reads or writes the ChildNodes property on the MSXML2 XMLElement compatibility object.

## Syntax
```asp
Dim obj, value
Set obj = Server.CreateObject("MSXML2.DOMDocument")
value = obj.ChildNodes
```

## Access
Read Only

## Return Values
Returns a Variant-compatible value.

## Remarks
- Child node collection.
- Property names are case-insensitive.

## Code Example
```asp
<%
Dim obj
Set obj = Server.CreateObject("MSXML2.DOMDocument")
On Error Resume Next
Response.Write CStr(obj.ChildNodes)
On Error GoTo 0
Set obj = Nothing
%>
```