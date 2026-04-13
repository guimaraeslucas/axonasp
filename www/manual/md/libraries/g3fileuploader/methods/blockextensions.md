# BlockExtensions Method

## Overview
Submits a comma-separated format string to define multiple blocked file extensions simultaneously.

## Syntax
```asp
Set uploader = Server.CreateObject("G3FILEUPLOADER")
uploader.BlockExtensions "jsp,asp,exe,dll,bat"
```

## Parameters and Arguments
- `ExtensionsList` (String, Required): A robust comma-separated blocklist of targeted formats.

## Return Values
Returns an `Empty` variant.
