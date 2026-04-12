# BrowserType.BackgroundSounds Property

## Overview
Reads or writes the BackgroundSounds member on the BrowserType compatibility object.

## Syntax
```asp
Dim obj, value
Set obj = Server.CreateObject("MSWC.BrowserType")
value = obj.BackgroundSounds
```

## Access
Read Only

## Return Values
Returns a Variant-compatible value.

## Remarks
- Indicates background sound capability.
- Property names are case-insensitive.

## Code Example
```asp
<%
Dim obj
Set obj = Server.CreateObject("MSWC.BrowserType")
Response.Write CStr(obj.BackgroundSounds)
Set obj = Nothing
%>
```