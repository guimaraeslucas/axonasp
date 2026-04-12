# BrowserType.MinorVer Property

## Overview
Reads or writes the MinorVer member on the BrowserType compatibility object.

## Syntax
```asp
Dim obj, value
Set obj = Server.CreateObject("MSWC.BrowserType")
value = obj.MinorVer
```

## Access
Read Only

## Return Values
Returns a Variant-compatible value.

## Remarks
- Detected minor version string.
- Property names are case-insensitive.

## Code Example
```asp
<%
Dim obj
Set obj = Server.CreateObject("MSWC.BrowserType")
Response.Write CStr(obj.MinorVer)
Set obj = Nothing
%>
```