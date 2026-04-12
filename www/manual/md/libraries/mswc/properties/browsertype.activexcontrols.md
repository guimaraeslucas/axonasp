# BrowserType.ActiveXControls Property

## Overview
Reads or writes the ActiveXControls member on the BrowserType compatibility object.

## Syntax
```asp
Dim obj, value
Set obj = Server.CreateObject("MSWC.BrowserType")
value = obj.ActiveXControls
```

## Access
Read Only

## Return Values
Returns a Variant-compatible value.

## Remarks
- Indicates ActiveX capability.
- Property names are case-insensitive.

## Code Example
```asp
<%
Dim obj
Set obj = Server.CreateObject("MSWC.BrowserType")
Response.Write CStr(obj.ActiveXControls)
Set obj = Nothing
%>
```