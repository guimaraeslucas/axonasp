# MyInfo.AnyCustomProperty Property

## Overview
Reads or writes the AnyCustomProperty member on the MyInfo compatibility object.

## Syntax
```asp
Dim obj, value
Set obj = Server.CreateObject("MSWC.MyInfo")
value = obj.AnyCustomProperty
```

## Access
Read Only

## Return Values
Returns a Variant-compatible value.

## Remarks
- Dynamic properties are loaded from MyInfo.xml element names.
- Property names are case-insensitive.

## Code Example
```asp
<%
Dim obj
Set obj = Server.CreateObject("MSWC.MyInfo")
Response.Write CStr(obj.AnyCustomProperty)
Set obj = Nothing
%>
```