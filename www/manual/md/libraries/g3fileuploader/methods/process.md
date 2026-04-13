# Process Method

## Overview
Processes a specific attached form file payload byte stream, pushing it directly to an allocated destination directory on disk, optionally allowing overrides of the core internal naming generator. Also supports the `Save` alias.

## Syntax
```asp
Set uploader = Server.CreateObject("G3FILEUPLOADER")
Dim res
Set res = uploader.Process("myFile", "/uploads/", "forcedName.png")
```

## Parameters and Arguments
- `FieldName` (String, Required): File form field identifier.
- `TargetDir` (String, Optional): Destination virtual directory to save the file (Defaults to "./").
- `NewFileName` (String, Optional): Explicit string name formatting option ignoring internal hash behaviors.

## Return Values
Returns a Dictionary object structured heavily reflecting successful operation indicators (via a boolean `IsSuccess` property). Other Dictionary keys include `OriginalFileName`, `NewFileName`, `Size`, `MimeType`, `Extension`, `FinalPath`, `RelativePath`, and `ErrorMessage`.
