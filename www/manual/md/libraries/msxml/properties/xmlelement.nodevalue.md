# XMLElement.NodeValue Property

## Overview
Reads or writes the NodeValue property on the MSXML2 XMLElement compatibility object.

## Syntax
```asp
Dim obj, value
Set obj = Server.CreateObject("MSXML2.DOMDocument")
value = obj.NodeValue
```

## Access
Read/Write

## Return Values
Returns a Variant-compatible value.

## Remarks
- Node value field.
- Property names are case-insensitive.

## Code Example
```asp
<%
Dim obj
Set obj = Server.CreateObject("MSXML2.DOMDocument")
On Error Resume Next
Response.Write CStr(obj.NodeValue)
On Error GoTo 0
Set obj = Nothing
%>
```