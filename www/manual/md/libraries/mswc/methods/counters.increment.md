# Counters.Increment Method

## Overview
Calls the Increment member on the Counters compatibility object.

## Syntax
```asp
Dim obj
Set obj = Server.CreateObject("MSWC.Counters")
n = obj.Increment(counterName)
```

## Parameters and Arguments
- Parameters are validated by runtime dispatch for MSWC.Counters.
- Invalid argument count or incompatible values can raise runtime errors.

## Return Values
Returns a Variant-compatible value according to operation semantics.

## Remarks
- Increments counter and returns new value.
- Use Set when assigning object return values.
- Member names are case-insensitive.

## Code Example
```asp
<%
Dim obj, result
Set obj = Server.CreateObject("MSWC.Counters")
result = Empty
On Error Resume Next
result = obj.Increment()
On Error GoTo 0
Response.Write CStr(result)
Set obj = Nothing
%>
```