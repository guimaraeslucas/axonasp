# MaxFileSize Property

## Overview
Gets or sets the maximum allowed size in bytes for an individual uploaded file.

## Syntax
```asp
uploader.MaxFileSize = sizeInBytes
sizeInBytes = uploader.MaxFileSize
```

## Parameters and Arguments
- `sizeInBytes` (Integer): The maximum file size in bytes.

## Return Values
Returns an **Integer** representing the size in bytes.

## Remarks
- The default value is 10,485,760 bytes (10 MB).
- Any file exceeding this limit will be rejected during the `Process` or `ProcessAll` methods.

## Code Example
```asp
<%
Dim uploader
Set uploader = Server.CreateObject("G3FILEUPLOADER")
' Set limit to 5 MB
uploader.MaxFileSize = 5242880
%>
```
