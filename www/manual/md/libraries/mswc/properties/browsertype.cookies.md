# BrowserType.Cookies Property

## Overview
Reads or writes the Cookies member on the BrowserType compatibility object.

## Syntax
```asp
Dim obj, value
Set obj = Server.CreateObject("MSWC.BrowserType")
value = obj.Cookies
```

## Access
Read Only

## Return Values
Returns a Variant-compatible value.

## Remarks
- Indicates cookies capability.
- Property names are case-insensitive.

## Code Example
```asp
<%
Dim obj
Set obj = Server.CreateObject("MSWC.BrowserType")
Response.Write CStr(obj.Cookies)
Set obj = Nothing
%>
```