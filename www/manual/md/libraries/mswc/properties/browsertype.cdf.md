# BrowserType.CDF Property

## Overview
Reads or writes the CDF member on the BrowserType compatibility object.

## Syntax
```asp
Dim obj, value
Set obj = Server.CreateObject("MSWC.BrowserType")
value = obj.CDF
```

## Access
Read Only

## Return Values
Returns a Variant-compatible value.

## Remarks
- Indicates CDF capability flag.
- Property names are case-insensitive.

## Code Example
```asp
<%
Dim obj
Set obj = Server.CreateObject("MSWC.BrowserType")
Response.Write CStr(obj.CDF)
Set obj = Nothing
%>
```