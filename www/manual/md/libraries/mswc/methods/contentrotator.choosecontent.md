# ContentRotator.ChooseContent Method

## Overview
Calls the ChooseContent member on the ContentRotator compatibility object.

## Syntax
```asp
Dim obj
Set obj = Server.CreateObject("MSWC.ContentRotator")
html = obj.ChooseContent(contentFile)
```

## Parameters and Arguments
- Parameters are validated by runtime dispatch for MSWC.ContentRotator.
- Invalid argument count or incompatible values can raise runtime errors.

## Return Values
Returns a Variant-compatible value according to operation semantics.

## Remarks
- Returns one weighted content block from rotator file.
- Use Set when assigning object return values.
- Member names are case-insensitive.

## Code Example
```asp
<%
Dim obj, result
Set obj = Server.CreateObject("MSWC.ContentRotator")
result = Empty
On Error Resume Next
result = obj.ChooseContent()
On Error GoTo 0
Response.Write CStr(result)
Set obj = Nothing
%>
```