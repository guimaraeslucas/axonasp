# PermissionChecker.HasAccess Method

## Overview
Calls the HasAccess member on the PermissionChecker compatibility object.

## Syntax
```asp
Dim obj
Set obj = Server.CreateObject("MSWC.PermissionChecker")
ok = obj.HasAccess(virtualPath)
```

## Parameters and Arguments
- Parameters are validated by runtime dispatch for MSWC.PermissionChecker.
- Invalid argument count or incompatible values can raise runtime errors.

## Return Values
Returns a Variant-compatible value according to operation semantics.

## Remarks
- Returns True when process can open mapped path for reading.
- Use Set when assigning object return values.
- Member names are case-insensitive.

## Code Example
```asp
<%
Dim obj, result
Set obj = Server.CreateObject("MSWC.PermissionChecker")
result = Empty
On Error Resume Next
result = obj.HasAccess()
On Error GoTo 0
Response.Write CStr(result)
Set obj = Nothing
%>
```