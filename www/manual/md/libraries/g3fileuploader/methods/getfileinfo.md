# GetFileInfo Method

## Overview
Retrieves metadata for a specific uploaded file field.

## Syntax
```asp
fileInfo = uploader.GetFileInfo(fieldName)
```

## Parameters and Arguments
- `fieldName` (String, Required): The name of the file input field.

## Return Values
Returns a **Dictionary** object containing file metadata, or **Empty** if the field was not found.

## Remarks
- The returned Dictionary includes keys: `OriginalFileName`, `Size`, `MimeType`, `Extension`, `IsValid`, and `ExceedsMaxSize`.
- Like `GetAllFilesInfo`, this does not perform any disk operations.

## Code Example
```asp
<%
Dim uploader, info
Set uploader = Server.CreateObject("G3FILEUPLOADER")
Set info = uploader.GetFileInfo("userAvatar")

If Not info Is Nothing Then
    Response.Write "Client Filename: " & info("OriginalFileName")
End If
%>
```
