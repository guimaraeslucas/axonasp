# BrowserType.Frames Property

## Overview
Reads or writes the Frames member on the BrowserType compatibility object.

## Syntax
```asp
Dim obj, value
Set obj = Server.CreateObject("MSWC.BrowserType")
value = obj.Frames
```

## Access
Read Only

## Return Values
Returns a Variant-compatible value.

## Remarks
- Indicates frames capability.
- Property names are case-insensitive.

## Code Example
```asp
<%
Dim obj
Set obj = Server.CreateObject("MSWC.BrowserType")
Response.Write CStr(obj.Frames)
Set obj = Nothing
%>
```