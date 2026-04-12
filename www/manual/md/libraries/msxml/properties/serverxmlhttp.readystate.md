# ServerXMLHTTP.ReadyState Property

## Overview
Reads or writes the ReadyState property on the MSXML2 ServerXMLHTTP compatibility object.

## Syntax
```asp
Dim obj, value
Set obj = Server.CreateObject("MSXML2.ServerXMLHTTP")
value = obj.ReadyState
```

## Access
Read Only

## Return Values
Returns a Variant-compatible value.

## Remarks
- Request state indicator.
- Property names are case-insensitive.

## Code Example
```asp
<%
Dim obj
Set obj = Server.CreateObject("MSXML2.ServerXMLHTTP")
On Error Resume Next
Response.Write CStr(obj.ReadyState)
On Error GoTo 0
Set obj = Nothing
%>
```