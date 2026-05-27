# Process Method

## Overview
Processes and saves a single uploaded file to a specified directory. Also supports the `Save` alias.

## Syntax
```asp
Set result = uploader.Process(fieldName, targetDir, newFileName)
```

## Parameters and Arguments
- `fieldName` (String, Required): The name of the file input field.
- `targetDir` (String, Optional): The destination directory. Defaults to the current directory ("./").
- `newFileName` (String, Optional): An optional name to rename the file. If omitted, a unique name is generated or the original name is used (based on `PreserveOriginalName`).

## Return Values
Returns a **Dictionary** object containing the operation result.
- `IsSuccess` (Boolean): **True** if saved successfully.
- `ErrorMessage` (String): Error description if `IsSuccess` is **False**.
- `FinalPath` (String): Absolute path on the server.
- `RelativePath` (String): Relative path from the web root.
- `OriginalFileName`, `NewFileName`, `Size`, `MimeType`, `Extension`.

## Remarks
- The `targetDir` is relative to the web root unless `AllowAbsolutePaths` is enabled.
- If `newFileName` does not contain an extension, the original extension is appended.

## Code Example
```asp
<%
Dim uploader, result
Set uploader = Server.CreateObject("G3FILEUPLOADER")
Set result = uploader.Process("fileInput", "/uploads/images", "profile_pic")

If result("IsSuccess") Then
    Response.Write "File saved as: " & result("NewFileName")
Else
    Response.Write "Upload failed: " & result("ErrorMessage")
End If
%>
```
