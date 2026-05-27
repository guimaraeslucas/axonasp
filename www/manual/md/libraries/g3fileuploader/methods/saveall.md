# SaveAll Method

## Overview
Alias for the `ProcessAll` method. Saves all uploaded files to disk.

## Syntax
```asp
results = uploader.SaveAll(targetDir)
```

## Parameters and Arguments
- `targetDir` (String, Optional): The destination directory.

## Return Values
Returns an **Array of Dictionary** result objects.

## Remarks
- Refer to the `ProcessAll` method documentation for detailed behavior.

## Code Example
```asp
<%
Dim uploader, results
Set uploader = Server.CreateObject("G3FILEUPLOADER")
results = uploader.SaveAll("temp/")
%>
```
