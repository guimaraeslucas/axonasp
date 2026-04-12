# BrowserType.MajorVer Property

## Overview
Reads or writes the MajorVer member on the BrowserType compatibility object.

## Syntax
```asp
Dim obj, value
Set obj = Server.CreateObject("MSWC.BrowserType")
value = obj.MajorVer
```

## Access
Read Only

## Return Values
Returns a Variant-compatible value.

## Remarks
- Detected major version string.
- Property names are case-insensitive.

## Code Example
```asp
<%
Dim obj
Set obj = Server.CreateObject("MSWC.BrowserType")
Response.Write CStr(obj.MajorVer)
Set obj = Nothing
%>
```