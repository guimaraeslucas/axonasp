# BrowserType.VBScript Property

## Overview
Reads or writes the VBScript member on the BrowserType compatibility object.

## Syntax
```asp
Dim obj, value
Set obj = Server.CreateObject("MSWC.BrowserType")
value = obj.VBScript
```

## Access
Read Only

## Return Values
Returns a Variant-compatible value.

## Remarks
- Indicates VBScript support capability.
- Property names are case-insensitive.

## Code Example
```asp
<%
Dim obj
Set obj = Server.CreateObject("MSWC.BrowserType")
Response.Write CStr(obj.VBScript)
Set obj = Nothing
%>
```