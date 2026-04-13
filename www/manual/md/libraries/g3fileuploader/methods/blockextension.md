# BlockExtension Method

## Overview
Registers a specific file extension to heavily restrict it. The system automatically rejects uploads containing this extension.

## Syntax
```asp
Set uploader = Server.CreateObject("G3FILEUPLOADER")
uploader.BlockExtension "php"
```

## Parameters and Arguments
- `Extension` (String, Required): The prohibited file format extension.

## Return Values
Returns an `Empty` variant.
