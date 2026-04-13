# Size Method

## Overview
Returns the size of a specified file in bytes.

## Syntax
```asp
fileSize = files.Size(path)
```

## Parameters and Arguments
- **path** (String, Required): The path to the file.

## Return Values
Returns an **Integer** (int64) representing the size of the file in bytes. If the file does not exist or is a directory, it returns 0.

## Remarks
- This method provides a quick way to check file metadata without reading the content.
- Path resolution is relative to the AxonASP sandbox root.

## Code Example
```asp
<%
Dim files, s
Set files = Server.CreateObject("G3FILES")
s = files.Size("/backups/dump.sql")
Response.Write "Backup size: " & FormatNumber(s / 1024, 2) & " KB"
Set files = Nothing
%>
```
