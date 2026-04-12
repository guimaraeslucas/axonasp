# AdRotator.TargetFrame Property

## Overview
Reads or writes the TargetFrame member on the AdRotator compatibility object.

## Syntax
```asp
Dim obj, value
Set obj = Server.CreateObject("MSWC.AdRotator")
value = obj.TargetFrame
```

## Access
Read/Write

## Return Values
Returns a Variant-compatible value.

## Remarks
- Sets optional target frame on generated anchor element.
- Property names are case-insensitive.

## Code Example
```asp
<%
Dim obj
Set obj = Server.CreateObject("MSWC.AdRotator")
Response.Write CStr(obj.TargetFrame)
Set obj = Nothing
%>
```