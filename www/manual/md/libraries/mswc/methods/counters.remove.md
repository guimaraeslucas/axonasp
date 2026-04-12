# Counters.Remove Method

## Overview
Calls the Remove member on the Counters compatibility object.

## Syntax
```asp
Dim obj
Set obj = Server.CreateObject("MSWC.Counters")
obj.Remove counterName
```

## Parameters and Arguments
- Parameters are validated by runtime dispatch for MSWC.Counters.
- Invalid argument count or incompatible values can raise runtime errors.

## Return Values
Returns a Variant-compatible value according to operation semantics.

## Remarks
- Deletes named counter entry.
- Use Set when assigning object return values.
- Member names are case-insensitive.

## Code Example
```asp
<%
Dim obj, result
Set obj = Server.CreateObject("MSWC.Counters")
result = Empty
On Error Resume Next
result = obj.Remove()
On Error GoTo 0
Response.Write CStr(result)
Set obj = Nothing
%>
```