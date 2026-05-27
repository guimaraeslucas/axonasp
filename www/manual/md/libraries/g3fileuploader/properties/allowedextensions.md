# AllowedExtensions Property

## Overview
Returns an array of file extensions that are explicitly allowed for upload.

## Syntax
```asp
extArray = uploader.AllowedExtensions
```

## Parameters and Arguments
None.

## Return Values
Returns an **Array of String** containing the allowed extensions.

## Remarks
- This property is read-only. Use the `AllowExtension` or `AllowExtensions` methods to modify this list.
- This list is only enforced if `SetUseAllowedOnly` is set to **True**.

## Code Example
```asp
<%
Dim uploader, exts, i
Set uploader = Server.CreateObject("G3FILEUPLOADER")
uploader.AllowExtensions "jpg, png, gif"

exts = uploader.AllowedExtensions
For i = 0 To UBound(exts)
    Response.Write "Allowed: " & exts(i) & "<br>"
Next
%>
```
