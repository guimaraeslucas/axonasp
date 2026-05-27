# GetAllFilesInfo Method

## Overview
Returns metadata for all files uploaded in the current multipart request.

## Syntax
```asp
infoArray = uploader.GetAllFilesInfo()
```

## Parameters and Arguments
None.

## Return Values
Returns an **Array of Dictionary** objects. Each Dictionary contains metadata for one file.

## Remarks
- Metadata includes `OriginalFileName`, `Size`, `MimeType`, `Extension`, `IsValid`, and `ExceedsMaxSize`.
- This method does not save files to disk; it only provides information about the request payload.

## Code Example
```asp
<%
Dim uploader, allFiles, fileInfo, i
Set uploader = Server.CreateObject("G3FILEUPLOADER")
allFiles = uploader.GetAllFilesInfo()

For i = 0 To UBound(allFiles)
    Set fileInfo = allFiles(i)
    Response.Write "File: " & fileInfo("OriginalFileName") & " (" & fileInfo("Size") & " bytes)<br>"
Next
%>
```
