# AddFolder Method

## Overview
Recursively includes a directory and all of its contents into the current write-mode archive in the G3Pix AxonASP G3ZIP library.

## Syntax
```asp
success = zip.AddFolder(sourcePath [, nameInZip])
```

## Parameters and Arguments
- **sourcePath** (String, Required): The path to the directory on the server to be added.
- **nameInZip** (String, Optional): The relative path the folder will have inside the ZIP archive. If omitted, the library uses the base name of the source directory.

## Return Values
Returns a **Boolean** indicating whether the folder and its contents were successfully added.

## Remarks
- The object must be in **Write** mode.
- This method performs a full recursive traversal of the specified directory.
- All files and subdirectories found will preserve their relative structure within the archive.

## Code Example
```asp
<%
Dim zip
Set zip = Server.CreateObject("G3ZIP")
If zip.Create("/backups/site_backup.zip") Then
    ' Add the entire 'content' folder to the root of the ZIP
    zip.AddFolder "/www/content", ""
    zip.Close
End If
Set zip = Nothing
%>
```
