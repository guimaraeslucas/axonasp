# ServerXMLHTTP.Timeout Property

## Overview
Reads or writes the Timeout property on the MSXML2 ServerXMLHTTP compatibility object.

## Syntax
```asp
Dim obj, value
Set obj = Server.CreateObject("MSXML2.ServerXMLHTTP")
value = obj.Timeout
```

## Access
Read/Write

## Return Values
Returns a Variant-compatible value.

## Remarks
- HTTP timeout in milliseconds.
- Property names are case-insensitive.

## Code Example
```asp
<%
Dim obj
Set obj = Server.CreateObject("MSXML2.ServerXMLHTTP")
On Error Resume Next
Response.Write CStr(obj.Timeout)
On Error GoTo 0
Set obj = Nothing
%>
```