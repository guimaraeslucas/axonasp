# ServerXMLHTTP.Open Method

## Overview
Calls the Open member on the MSXML2 ServerXMLHTTP compatibility object.

## Syntax
```asp
Dim obj
Set obj = Server.CreateObject("MSXML2.ServerXMLHTTP")
obj.Open "GET", url, False
```

## Parameters and Arguments
- Parameters are validated by runtime dispatch for this object.
- Invalid argument count or incompatible values can raise runtime errors.

## Return Values
Returns a Variant-compatible value or native object handle depending on the operation.

## Remarks
- Configures HTTP method, URL, and async flag.
- Member names are case-insensitive.
- Use Set for object return values.

## Code Example
```asp
<%
Dim obj
Set obj = Server.CreateObject("MSXML2.ServerXMLHTTP")
On Error Resume Next
obj.Open
If Err.Number <> 0 Then
    Response.Write "Error: " & Err.Description
    Err.Clear
End If
On Error GoTo 0
Set obj = Nothing
%>
```