# BrowserType.Version Property

## Overview
Reads or writes the Version member on the BrowserType compatibility object.

## Syntax
```asp
Dim obj, value
Set obj = Server.CreateObject("MSWC.BrowserType")
value = obj.Version
```

## Access
Read Only

## Return Values
Returns a Variant-compatible value.

## Remarks
- Detected browser version string.
- Property names are case-insensitive.

## Code Example
```asp
<%
Dim obj
Set obj = Server.CreateObject("MSWC.BrowserType")
Response.Write CStr(obj.Version)
Set obj = Nothing
%>
```