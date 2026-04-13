# GetFileInfo Method

## Overview
Extracts data related to a specified inbound form file input before you commit the asset onto system storage via the primary methods. 

## Syntax
```asp
Set uploader = Server.CreateObject("G3FILEUPLOADER")
Dim details
Set details = uploader.GetFileInfo("document")
```

## Parameters and Arguments
- `FieldName` (String, Required): The exact string moniker denoting the form input parameter mapping to a physical attachment stream.

## Return Values
Returns a Dictionary object populated mapping string configurations, or a Dictionary representing failure (with `IsSuccess` set to false) if missing details occur, or `Empty` if no HTTP request context is completely available.
