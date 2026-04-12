# ServerXMLHTTP.ResponseBody Property

## Overview
Reads or writes the ResponseBody property on the MSXML2 ServerXMLHTTP compatibility object.

## Syntax
```asp
Dim obj, value
Set obj = Server.CreateObject("MSXML2.ServerXMLHTTP")
value = obj.ResponseBody
```

## Access
Read Only

## Return Values
Returns a Variant-compatible value.

## Remarks
- Response body as byte array representation.
- Property names are case-insensitive.

## Code Example
```asp
<%
Dim obj
Set obj = Server.CreateObject("MSXML2.ServerXMLHTTP")
On Error Resume Next
Response.Write CStr(obj.ResponseBody)
On Error GoTo 0
Set obj = Nothing
%>
```