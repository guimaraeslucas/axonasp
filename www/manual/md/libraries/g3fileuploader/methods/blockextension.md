# BlockExtension Method

## Overview
Adds a single file extension to the blacklist of forbidden upload types.

## Syntax
```asp
uploader.BlockExtension extension
```

## Parameters and Arguments
- `extension` (String, Required): The file extension to block (e.g., "exe" or ".bat").

## Return Values
Returns **Empty**.

## Remarks
- Blocked extensions take precedence over allowed extensions.
- This is useful for preventing the upload of potentially malicious executable files.

## Code Example
```asp
<%
Dim uploader
Set uploader = Server.CreateObject("G3FILEUPLOADER")
uploader.BlockExtension "exe"
uploader.BlockExtension "msi"
%>
```
