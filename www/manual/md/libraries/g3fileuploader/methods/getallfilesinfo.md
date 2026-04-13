# GetAllFilesInfo Method

## Overview
Performs an architectural scan against the intercepted HTTP payload, surfacing comprehensive data dictionaries representing every uploaded item, bypassing persistence entirely. 

## Syntax
```asp
Set uploader = Server.CreateObject("G3FILEUPLOADER")
Dim fileArray
fileArray = uploader.GetAllFilesInfo()
```

## Return Values
Returns a VBScript array filled with Dictionary objects. Each Dictionary contains attributes analogous to `GetFileInfo`. If no multipart form data exists, returns an empty VBScript array.
