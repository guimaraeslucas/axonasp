# PreserveOriginalName Property

## Overview
Instructs the uploader engine to bypass its internal random hash generator and save files using their exact original names as submitted by the client device.

## Syntax
```asp
Set uploader = Server.CreateObject("G3FILEUPLOADER")
uploader.PreserveOriginalName = True
```

## Return Values
Sets or returns a boolean reflecting the name preservation rule.
