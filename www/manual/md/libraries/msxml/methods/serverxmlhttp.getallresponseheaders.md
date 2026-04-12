# ServerXMLHTTP.GetAllResponseHeaders Method

## Overview
Calls the GetAllResponseHeaders member on the MSXML2 ServerXMLHTTP compatibility object.

## Syntax
```asp
Dim obj
Set obj = Server.CreateObject("MSXML2.ServerXMLHTTP")
v = obj.GetAllResponseHeaders()
```

## Parameters and Arguments
- Parameters are validated by runtime dispatch for this object.
- Invalid argument count or incompatible values can raise runtime errors.

## Return Values
Returns a Variant-compatible value or native object handle depending on the operation.

## Remarks
- Returns CRLF-joined response headers text.
- Member names are case-insensitive.
- Use Set for object return values.

## Code Example
```asp
<%
Dim obj
Set obj = Server.CreateObject("MSXML2.ServerXMLHTTP")
On Error Resume Next
obj.GetAllResponseHeaders
If Err.Number <> 0 Then
    Response.Write "Error: " & Err.Description
    Err.Clear
End If
On Error GoTo 0
Set obj = Nothing
%>
```