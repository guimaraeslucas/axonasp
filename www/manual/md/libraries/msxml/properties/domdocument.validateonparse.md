# DOMDocument.ValidateOnParse Property

## Overview
Reads or writes the ValidateOnParse property on the MSXML2 DOMDocument compatibility object.

## Syntax
```asp
Dim obj, value
Set obj = Server.CreateObject("MSXML2.DOMDocument")
value = obj.ValidateOnParse
```

## Access
Read/Write

## Return Values
Returns a Variant-compatible value.

## Remarks
- Compatibility validate-on-parse flag.
- Property names are case-insensitive.

## Code Example
```asp
<%
Dim obj
Set obj = Server.CreateObject("MSXML2.DOMDocument")
On Error Resume Next
Response.Write CStr(obj.ValidateOnParse)
On Error GoTo 0
Set obj = Nothing
%>
```