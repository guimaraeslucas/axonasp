# ServerXMLHTTP.Status Property

## Overview
Reads or writes the Status property on the MSXML2 ServerXMLHTTP compatibility object.

## Syntax
```asp
Dim obj, value
Set obj = Server.CreateObject("MSXML2.ServerXMLHTTP")
value = obj.Status
```

## Access
Read Only

## Return Values
Returns a Variant-compatible value.

## Remarks
- HTTP status code.
- Property names are case-insensitive.

## Code Example
```asp
<%
Dim obj
Set obj = Server.CreateObject("MSXML2.ServerXMLHTTP")
On Error Resume Next
Response.Write CStr(obj.Status)
On Error GoTo 0
Set obj = Nothing
%>
```