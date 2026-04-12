# BrowserType.JavaApplets Property

## Overview
Reads or writes the JavaApplets member on the BrowserType compatibility object.

## Syntax
```asp
Dim obj, value
Set obj = Server.CreateObject("MSWC.BrowserType")
value = obj.JavaApplets
```

## Access
Read Only

## Return Values
Returns a Variant-compatible value.

## Remarks
- Indicates Java applet capability.
- Property names are case-insensitive.

## Code Example
```asp
<%
Dim obj
Set obj = Server.CreateObject("MSWC.BrowserType")
Response.Write CStr(obj.JavaApplets)
Set obj = Nothing
%>
```