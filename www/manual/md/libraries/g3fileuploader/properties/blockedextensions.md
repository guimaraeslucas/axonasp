# BlockedExtensions Property

## Overview
Returns an array of file extensions that are blacklisted and will be rejected.

## Syntax
```asp
extArray = uploader.BlockedExtensions
```

## Parameters and Arguments
None.

## Return Values
Returns an **Array of String** containing the blocked extensions.

## Remarks
- This property is read-only. Use the `BlockExtension` or `BlockExtensions` methods to modify this list.
- Blocked extensions are always rejected, even if they appear in the `AllowedExtensions` list.

## Code Example
```asp
<%
Dim uploader, exts, i
Set uploader = Server.CreateObject("G3FILEUPLOADER")
uploader.BlockExtensions "exe, bat, cmd"

exts = uploader.BlockedExtensions
For i = 0 To UBound(exts)
    Response.Write "Blocked: " & exts(i) & "<br>"
Next
%>
```
