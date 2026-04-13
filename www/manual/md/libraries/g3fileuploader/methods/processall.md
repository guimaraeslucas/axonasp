# ProcessAll Method

## Overview
Saves all pending multipart files embedded directly onto the stream request simultaneously parsing the system configuration. Validates limits, unique variables, and dynamically writes to the filesystem tracking the final status for each target. Also accepts the `SaveAll` alias.

## Syntax
```asp
Set uploader = Server.CreateObject("G3FILEUPLOADER")
Dim resultsList
resultsList = uploader.ProcessAll("/uploads/")
```

## Parameters and Arguments
- `TargetDir` (String, Optional): Destination virtual directory to securely organize distribution of all incoming items (Defaults to "./").

## Return Values
Returns a VBScript array containing separate Dictionary objects mapping independent results for each item processed (sharing attributes akin to `Process`), such as `IsSuccess` and related naming metrics to capture standard logs precisely.
