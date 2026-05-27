# FormFields Property

## Overview
Returns a VBScript Dictionary containing all non-file form fields submitted in the multipart request.

## Syntax
```asp
Set fields = uploader.FormFields
```

## Parameters and Arguments
None.

## Return Values
Returns a **Dictionary** object where keys are field names and values are the corresponding string data.

## Remarks
- This property is read-only.
- It mimics the behavior of an associative array (similar to PHP's `$_POST`) and provides a convenient way to iterate over all received text inputs without knowing their names beforehand.
- Because Classic ASP cannot mix `Request.Form` and `Request.BinaryRead`, this property is the recommended way to access form data during a file upload.

## Code Example
```asp
<%
Dim uploader, fields, key
Set uploader = Server.CreateObject("G3FILEUPLOADER")
Set fields = uploader.FormFields

Response.Write "Submitted Fields:<br>"
For Each key In fields.Keys
    Response.Write key & ": " & fields(key) & "<br>"
Next
%>
```
