# DebugMode Property

## Overview
Gets or sets a Boolean value that enables or disables verbose debugging output for the uploader operations.

## Syntax
```asp
uploader.DebugMode = mode
mode = uploader.DebugMode
```

## Parameters and Arguments
- `mode` (Boolean): Set to **True** to enable debug logging, or **False** to disable it.

## Return Values
Returns a **Boolean** value.

## Remarks
- When enabled, additional operation details may be logged to the server console or log files to assist in troubleshooting upload issues.

## Code Example
```asp
<%
Dim uploader
Set uploader = Server.CreateObject("G3FILEUPLOADER")
uploader.DebugMode = True
%>
```
