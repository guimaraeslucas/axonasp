# ContentRotator.GetAllContent Method

## Overview
Calls the GetAllContent member on the ContentRotator compatibility object.

## Syntax
```asp
Dim obj
Set obj = Server.CreateObject("MSWC.ContentRotator")
html = obj.GetAllContent(contentFile)
```

## Parameters and Arguments
- Parameters are validated by runtime dispatch for MSWC.ContentRotator.
- Invalid argument count or incompatible values can raise runtime errors.

## Return Values
Returns a Variant-compatible value according to operation semantics.

## Remarks
- Returns all content blocks joined with separators.
- Use Set when assigning object return values.
- Member names are case-insensitive.

## Code Example
```asp
<%
Dim obj, result
Set obj = Server.CreateObject("MSWC.ContentRotator")
result = Empty
On Error Resume Next
result = obj.GetAllContent()
On Error GoTo 0
Response.Write CStr(result)
Set obj = Nothing
%>
```