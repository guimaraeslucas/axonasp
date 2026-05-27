# Save Method

## Overview
Alias for the `Process` method. Saves a single uploaded file.

## Syntax
```asp
Set result = uploader.Save(fieldName, targetDir, newFileName)
```

## Parameters and Arguments
- `fieldName` (String, Required): The name of the file input field.
- `targetDir` (String, Optional): Destination directory.
- `newFileName` (String, Optional): Optional new filename.

## Return Values
Returns a **Dictionary** result.

## Remarks
- Functionally identical to `Process`. Refer to the `Process` method documentation for details.

## Code Example
```asp
<%
Dim uploader, res
Set uploader = Server.CreateObject("G3FILEUPLOADER")
Set res = uploader.Save("myFile", "temp/", "")
%>
```
