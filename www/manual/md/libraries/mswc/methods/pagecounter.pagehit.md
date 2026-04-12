# PageCounter.PageHit Method

## Overview
Calls the PageHit member on the PageCounter compatibility object.

## Syntax
```asp
Dim obj
Set obj = Server.CreateObject("MSWC.PageCounter")
n = obj.PageHit()
```

## Parameters and Arguments
- Parameters are validated by runtime dispatch for MSWC.PageCounter.
- Invalid argument count or incompatible values can raise runtime errors.

## Return Values
Returns a Variant-compatible value according to operation semantics.

## Remarks
- Increments and returns current page hit count.
- Use Set when assigning object return values.
- Member names are case-insensitive.

## Code Example
```asp
<%
Dim obj, result
Set obj = Server.CreateObject("MSWC.PageCounter")
result = Empty
On Error Resume Next
result = obj.PageHit()
On Error GoTo 0
Response.Write CStr(result)
Set obj = Nothing
%>
```