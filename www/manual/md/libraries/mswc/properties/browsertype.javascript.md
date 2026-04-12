# BrowserType.JavaScript Property

## Overview
Reads or writes the JavaScript member on the BrowserType compatibility object.

## Syntax
```asp
Dim obj, value
Set obj = Server.CreateObject("MSWC.BrowserType")
value = obj.JavaScript
```

## Access
Read Only

## Return Values
Returns a Variant-compatible value.

## Remarks
- Indicates JavaScript capability.
- Property names are case-insensitive.

## Code Example
```asp
<%
Dim obj
Set obj = Server.CreateObject("MSWC.BrowserType")
Response.Write CStr(obj.JavaScript)
Set obj = Nothing
%>
```