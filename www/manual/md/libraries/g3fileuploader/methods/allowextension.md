# AllowExtension Method

## Overview
Appends a single file extension to the internally maintained allowed extensions permit list. Usually used in tandem with the `SetUseAllowedOnly` mechanism.

## Syntax
```asp
Set uploader = Server.CreateObject("G3FILEUPLOADER")
uploader.AllowExtension ".pdf"
```

## Parameters and Arguments
- `Extension` (String, Required): The literal file extension (with or without the leading dot).

## Return Values
Returns an `Empty` variant.
