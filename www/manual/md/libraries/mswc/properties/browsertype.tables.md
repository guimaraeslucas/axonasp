# BrowserType.Tables Property

## Overview
Reads or writes the Tables member on the BrowserType compatibility object.

## Syntax
```asp
Dim obj, value
Set obj = Server.CreateObject("MSWC.BrowserType")
value = obj.Tables
```

## Access
Read Only

## Return Values
Returns a Variant-compatible value.

## Remarks
- Indicates tables capability.
- Property names are case-insensitive.

## Code Example
```asp
<%
Dim obj
Set obj = Server.CreateObject("MSWC.BrowserType")
Response.Write CStr(obj.Tables)
Set obj = Nothing
%>
```