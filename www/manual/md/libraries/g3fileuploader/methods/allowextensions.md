# AllowExtensions Method

## Overview
Adds multiple file extensions to the whitelist from a comma-separated string.

## Syntax
```asp
uploader.AllowExtensions extensions
```

## Parameters and Arguments
- `extensions` (String, Required): A comma-separated list of extensions (e.g., "jpg, png, gif").

## Return Values
Returns **Empty**.

## Remarks
- This is a convenience method for calling `AllowExtension` multiple times.
- If `SetUseAllowedOnly` is **True**, the uploader restricts all files to this list.

## Code Example
```asp
<%
Dim uploader
Set uploader = Server.CreateObject("G3FILEUPLOADER")
uploader.AllowExtensions "docx, xlsx, pptx, pdf"
%>
```
