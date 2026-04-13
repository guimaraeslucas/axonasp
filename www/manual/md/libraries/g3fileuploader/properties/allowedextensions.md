# AllowedExtensions Property

## Overview
A readonly property that returns an array of explicitly permitted file extensions currently loaded into memory.

## Syntax
```asp
Set uploader = Server.CreateObject("G3FILEUPLOADER")
Dim list
list = uploader.AllowedExtensions
```

## Return Values
Returns a VBScript array containing strings of all allowed file extensions.
