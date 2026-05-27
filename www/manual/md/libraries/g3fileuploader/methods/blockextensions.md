# BlockExtensions Method

## Overview
Adds multiple file extensions to the blacklist from a comma-separated string.

## Syntax
```asp
uploader.BlockExtensions extensions
```

## Parameters and Arguments
- `extensions` (String, Required): A comma-separated list of extensions to block.

## Return Values
Returns **Empty**.

## Remarks
- Convenience method for blocking multiple file types in a single call.
- Any file with an extension in this list will be rejected during the `Process` or `ProcessAll` methods.

## Code Example
```asp
<%
Dim uploader
Set uploader = Server.CreateObject("G3FILEUPLOADER")
uploader.BlockExtensions "php, asp, aspx, jsp, exe, bat"
%>
```
