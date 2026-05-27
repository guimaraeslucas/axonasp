# AllowExtension Method

## Overview
Adds a single file extension to the whitelist of allowed upload types.

## Syntax
```asp
uploader.AllowExtension extension
```

## Parameters and Arguments
- `extension` (String, Required): The file extension to allow (e.g., "jpg" or ".pdf"). Leading dots are optional.

## Return Values
Returns **Empty**.

## Remarks
- If `SetUseAllowedOnly` is set to **True**, only extensions added via `AllowExtension` or `AllowExtensions` will be accepted.
- Input is case-insensitive and leading spaces are trimmed.

## Code Example
```asp
<%
Dim uploader
Set uploader = Server.CreateObject("G3FILEUPLOADER")
uploader.AllowExtension "jpg"
uploader.AllowExtension ".png"
%>
```
