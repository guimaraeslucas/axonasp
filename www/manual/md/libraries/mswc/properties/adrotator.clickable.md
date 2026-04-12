# AdRotator.Clickable Property

## Overview
Reads or writes the Clickable member on the AdRotator compatibility object.

## Syntax
```asp
Dim obj, value
Set obj = Server.CreateObject("MSWC.AdRotator")
value = obj.Clickable
```

## Access
Read/Write

## Return Values
Returns a Variant-compatible value.

## Remarks
- Controls if generated output wraps IMG in anchor tag.
- Property names are case-insensitive.

## Code Example
```asp
<%
Dim obj
Set obj = Server.CreateObject("MSWC.AdRotator")
Response.Write CStr(obj.Clickable)
Set obj = Nothing
%>
```