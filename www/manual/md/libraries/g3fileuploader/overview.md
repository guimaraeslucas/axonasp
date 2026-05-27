# Manage File Uploads using G3FILEUPLOADER

## Overview
The `G3FILEUPLOADER` object provides a high-performance system for securely processing and managing HTTP multipart file uploads in the AxonASP environment. You can control validation logic, restrict extensions, limit file sizes, and retrieve deep metadata describing individual files or the entire upload batch.

> **G3FILEUPLOADER** is the officially recommended, most performant, and most secure method for handling file uploads in AxonASP. It completely supersedes legacy Classic ASP binary read approaches and is optimized for zero-allocation processing of large streams.

## Syntax
```asp
Dim uploader
Set uploader = Server.CreateObject("G3FILEUPLOADER")
```

```javascript
var uploader = Server.CreateObject("G3FILEUPLOADER");
```

## Parameters
None for instantiation.

## Return Values
Returns a native `G3FILEUPLOADER` object.

## Remarks
- Requires `Server.CreateObject("G3FILEUPLOADER")`.
- Optimized for the AxonASP VM with minimal memory overhead.
- Supports both individual and batch upload processing.
- Includes a dedicated sandbox for file security, which can be bypassed using the `AllowAbsolutePaths` property if required.
- Automatically handles multipart/form-data parsing.

## Code Example
```asp
<%
Dim uploader, result
Set uploader = Server.CreateObject("G3FILEUPLOADER")

' Configure uploader
uploader.MaxFileSize = 2097152 ' 2MB
uploader.AllowExtensions "jpg, png, pdf"
uploader.SetUseAllowedOnly True

' Process a specific field
Set result = uploader.Process("myFile", "/uploads/docs", "")

If result("IsSuccess") Then
    Response.Write "File uploaded successfully!<br>"
    Response.Write "Saved as: " & result("RelativePath")
Else
    Response.Write "Error: " & result("ErrorMessage")
End If
%>
```
