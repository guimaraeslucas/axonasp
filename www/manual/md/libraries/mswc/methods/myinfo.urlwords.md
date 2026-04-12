# MyInfo.URLWords Method

## Overview
Calls the URLWords member on the MyInfo compatibility object.

## Syntax
```asp
Dim obj
Set obj = Server.CreateObject("MSWC.MyInfo")
value = obj.URLWords(index)
```

## Parameters and Arguments
- Parameters are validated by runtime dispatch for MSWC.MyInfo.
- Invalid argument count or incompatible values can raise runtime errors.

## Return Values
Returns a Variant-compatible value according to operation semantics.

## Remarks
- Returns URLWords field value by suffix index from MyInfo.xml.
- Use Set when assigning object return values.
- Member names are case-insensitive.

## Code Example
```asp
<%
Dim obj, result
Set obj = Server.CreateObject("MSWC.MyInfo")
result = Empty
On Error Resume Next
result = obj.URLWords()
On Error GoTo 0
Response.Write CStr(result)
Set obj = Nothing
%>
```