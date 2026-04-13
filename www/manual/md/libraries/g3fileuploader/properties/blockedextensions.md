# BlockedExtensions Property

## Overview
A readonly property that returns an array of explicitly forbidden file extensions loaded into the validation engine. 

## Syntax
```asp
Set uploader = Server.CreateObject("G3FILEUPLOADER")
Dim list
list = uploader.BlockedExtensions
```

## Return Values
Returns a VBScript array containing strings of all strictly blocked file extensions.
