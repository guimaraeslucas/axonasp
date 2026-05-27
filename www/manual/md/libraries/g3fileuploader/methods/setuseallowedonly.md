# SetUseAllowedOnly Method

## Overview
Configures whether the uploader should strictly allow only files with extensions explicitly present in the `AllowedExtensions` list.

## Syntax
```asp
uploader.SetUseAllowedOnly mode
```

## Parameters and Arguments
- `mode` (Boolean, Required): Set to **True** to enable strict allow-list mode, or **False** to allow all files except those in the `BlockedExtensions` list.

## Return Values
Returns **Empty**.

## Remarks
- By default, this mode is **False**, meaning all extensions are allowed unless they are specifically blocked.
- When enabled, you must populate the `AllowedExtensions` list using `AllowExtension` or `AllowExtensions`.

## Code Example
```asp
<%
Dim uploader
Set uploader = Server.CreateObject("G3FILEUPLOADER")
uploader.AllowExtensions "jpg, png"
uploader.SetUseAllowedOnly True

' Only JPG and PNG will be accepted now.
%>
```
