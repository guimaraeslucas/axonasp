# AllowExtensions Method

## Overview
Bulk injects multiple allowed file extensions by processing a comma-delimited string of formats.

## Syntax
```asp
Set uploader = Server.CreateObject("G3FILEUPLOADER")
uploader.AllowExtensions "pdf,doc,docx"
```

## Parameters and Arguments
- `ExtensionsList` (String, Required): A comma-separated list of file extensions.

## Return Values
Returns an `Empty` variant.
