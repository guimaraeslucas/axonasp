# Form Method

## Overview
Retrieves the value of a non-file form field from a multipart/form-data request.

## Syntax
```asp
value = uploader.Form(fieldName)
```

## Parameters and Arguments
- `fieldName` (String, Required): The name of the form input field.

## Return Values
Returns a **String** containing the field value, or **Empty** if the field was not found or the request is not multipart.

## Remarks
- This method is an alternative to `Request.Form` when handling file uploads, as standard binary reads and form collection access cannot be mixed in Classic ASP.
- If multiple values exist for the same field name, only the first one is returned.
- Also supports the `FormValue` alias.

## Code Example
```asp
<%
Dim uploader, userName
Set uploader = Server.CreateObject("G3FILEUPLOADER")
userName = uploader.Form("txtName")

If userName <> "" Then
    Response.Write "Hello, " & userName
End If
%>
```
