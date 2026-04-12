# Tools.FileExists Method

## Overview
Calls the FileExists member on the Tools compatibility object.

## Syntax
```asp
Dim obj
Set obj = Server.CreateObject("MSWC.Tools")
ok = obj.FileExists(virtualPath)
```

## Parameters and Arguments
- Parameters are validated by runtime dispatch for MSWC.Tools.
- Invalid argument count or incompatible values can raise runtime errors.

## Return Values
Returns a Variant-compatible value according to operation semantics.

## Remarks
- Returns True when mapped file exists.
- Use Set when assigning object return values.
- Member names are case-insensitive.

## Code Example
```asp
<%
Dim obj, result
Set obj = Server.CreateObject("MSWC.Tools")
result = Empty
On Error Resume Next
result = obj.FileExists()
On Error GoTo 0
Response.Write CStr(result)
Set obj = Nothing
%>
```