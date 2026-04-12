# DOMDocument.ResolveExternals Property

## Overview
Reads or writes the ResolveExternals property on the MSXML2 DOMDocument compatibility object.

## Syntax
```asp
Dim obj, value
Set obj = Server.CreateObject("MSXML2.DOMDocument")
value = obj.ResolveExternals
```

## Access
Read/Write

## Return Values
Returns a Variant-compatible value.

## Remarks
- Compatibility resolve externals flag.
- Property names are case-insensitive.

## Code Example
```asp
<%
Dim obj
Set obj = Server.CreateObject("MSXML2.DOMDocument")
On Error Resume Next
Response.Write CStr(obj.ResolveExternals)
On Error GoTo 0
Set obj = Nothing
%>
```