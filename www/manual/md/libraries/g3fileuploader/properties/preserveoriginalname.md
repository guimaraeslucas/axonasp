# PreserveOriginalName Property

## Overview
Gets or sets a Boolean value that determines whether the uploader should keep the client's original filename when saving to disk.

## Syntax
```asp
uploader.PreserveOriginalName = mode
mode = uploader.PreserveOriginalName
```

## Parameters and Arguments
- `mode` (Boolean): Set to **True** to keep the original filename, or **False** to generate a unique random filename.

## Return Values
Returns a **Boolean** value.

## Remarks
- The default value is **False** (random filename generation) to prevent filename collisions and security risks associated with malicious filenames.
- Even if set to **True**, you can still provide an explicit filename in the `Process` method's third argument.

## Code Example
```asp
<%
Dim uploader
Set uploader = Server.CreateObject("G3FILEUPLOADER")
uploader.PreserveOriginalName = True
%>
```
