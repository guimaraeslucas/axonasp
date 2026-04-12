# BrowserType.Browser Property

## Overview
Reads or writes the Browser member on the BrowserType compatibility object.

## Syntax
```asp
Dim obj, value
Set obj = Server.CreateObject("MSWC.BrowserType")
value = obj.Browser
```

## Access
Read Only

## Return Values
Returns a Variant-compatible value.

## Remarks
- Detected browser family name.
- Property names are case-insensitive.

## Code Example
```asp
<%
Dim obj
Set obj = Server.CreateObject("MSWC.BrowserType")
Response.Write CStr(obj.Browser)
Set obj = Nothing
%>
```