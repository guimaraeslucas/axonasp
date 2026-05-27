# AllowAbsolutePaths Property

## Overview
Gets or sets a Boolean value that determines whether the uploader is allowed to save files to absolute system paths.

## Syntax
```asp
uploader.AllowAbsolutePaths = mode
mode = uploader.AllowAbsolutePaths
```

## Parameters and Arguments
- `mode` (Boolean): Set to **True** to allow saving files to absolute paths (e.g., "C:\uploads" or "/var/www/uploads"). Default is **False**.

## Return Values
Returns a **Boolean** value.

## Remarks
- By default, `G3FILEUPLOADER` restricts all file operations to the web root sandbox for security.
- Enabling this property allows the application to save files anywhere on the host system where the AxonASP process has write permissions.
- **SECURITY NOTICE:** Use this property with extreme caution and ensure that destination paths are not derived from untrusted user input.

## Code Example
```asp
<%
Dim uploader, result
Set uploader = Server.CreateObject("G3FILEUPLOADER")
uploader.AllowAbsolutePaths = True

' Saving to an absolute system path
Set result = uploader.Process("file1", "C:\MySystemUploads\", "")
%>
```
