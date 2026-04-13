# MaxFileSize Property

## Overview
Controls the maximum permitted file size (in bytes) per individual request. By default, this is set to 10 MB (10485760 bytes).

## Syntax
```asp
Set uploader = Server.CreateObject("G3FILEUPLOADER")
uploader.MaxFileSize = 5242880 ' 5 MB maximum
```

## Return Values
Sets or returns a long integer reflecting the maximum file size configured in the uploader instance.
