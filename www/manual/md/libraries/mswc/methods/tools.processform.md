# Tools.ProcessForm Method

## Overview
Calls the ProcessForm member on the Tools compatibility object.

## Syntax
```asp
Dim obj
Set obj = Server.CreateObject("MSWC.Tools")
obj.ProcessForm
```

## Parameters and Arguments
- Parameters are validated by runtime dispatch for MSWC.Tools.
- Invalid argument count or incompatible values can raise runtime errors.

## Return Values
Returns a Variant-compatible value according to operation semantics.

## Remarks
- Compatibility stub; currently no-op.
- Use Set when assigning object return values.
- Member names are case-insensitive.

## Code Example
```asp
<%
Dim obj, result
Set obj = Server.CreateObject("MSWC.Tools")
result = Empty
On Error Resume Next
result = obj.ProcessForm()
On Error GoTo 0
Response.Write CStr(result)
Set obj = Nothing
%>
```