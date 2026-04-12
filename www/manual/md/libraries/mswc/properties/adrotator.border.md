# AdRotator.Border Property

## Overview
Reads or writes the Border member on the AdRotator compatibility object.

## Syntax
```asp
Dim obj, value
Set obj = Server.CreateObject("MSWC.AdRotator")
value = obj.Border
```

## Access
Read/Write

## Return Values
Returns a Variant-compatible value.

## Remarks
- Overrides border value used by generated IMG tag.
- Property names are case-insensitive.

## Code Example
```asp
<%
Dim obj
Set obj = Server.CreateObject("MSWC.AdRotator")
Response.Write CStr(obj.Border)
Set obj = Nothing
%>
```