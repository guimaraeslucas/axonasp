# NextLink.GetListIndex Method

## Overview
Calls the GetListIndex member on the NextLink compatibility object.

## Syntax
```asp
Dim obj
Set obj = Server.CreateObject("MSWC.NextLink")
n = obj.GetListIndex(listFile)
```

## Parameters and Arguments
- Parameters are validated by runtime dispatch for MSWC.NextLink.
- Invalid argument count or incompatible values can raise runtime errors.

## Return Values
Returns a Variant-compatible value according to operation semantics.

## Remarks
- Returns current item index based on current request URL.
- Use Set when assigning object return values.
- Member names are case-insensitive.

## Code Example
```asp
<%
Dim obj, result
Set obj = Server.CreateObject("MSWC.NextLink")
result = Empty
On Error Resume Next
result = obj.GetListIndex()
On Error GoTo 0
Response.Write CStr(result)
Set obj = Nothing
%>
```