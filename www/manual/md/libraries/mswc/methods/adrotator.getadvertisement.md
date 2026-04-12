# AdRotator.GetAdvertisement Method

## Overview
Calls the GetAdvertisement member on the AdRotator compatibility object.

## Syntax
```asp
Dim obj
Set obj = Server.CreateObject("MSWC.AdRotator")
html = obj.GetAdvertisement(scheduleFile)
```

## Parameters and Arguments
- Parameters are validated by runtime dispatch for MSWC.AdRotator.
- Invalid argument count or incompatible values can raise runtime errors.

## Return Values
Returns a Variant-compatible value according to operation semantics.

## Remarks
- Reads an ad schedule file and returns the selected ad HTML fragment.
- Use Set when assigning object return values.
- Member names are case-insensitive.

## Code Example
```asp
<%
Dim obj, result
Set obj = Server.CreateObject("MSWC.AdRotator")
result = Empty
On Error Resume Next
result = obj.GetAdvertisement()
On Error GoTo 0
Response.Write CStr(result)
Set obj = Nothing
%>
```