# Copy Method

## Overview
Copies a file from a source path to a destination path.

## Syntax
```asp
success = files.Copy(source, dest)
```

## Parameters and Arguments
- **source** (String, Required): The path to the file to be copied.
- **dest** (String, Required): The target path for the copy.

## Return Values
Returns a **Boolean** indicating whether the copy operation was successful.

## Remarks
- If the destination file already exists, it will be overwritten.
- Both paths are relative to the AxonASP sandbox root.

## Code Example
```asp
<%
Dim files
Set files = Server.CreateObject("G3FILES")
If files.Copy("/data/source.txt", "/backups/copy.txt") Then
    Response.Write "File copied."
End If
Set files = Nothing
%>
```
