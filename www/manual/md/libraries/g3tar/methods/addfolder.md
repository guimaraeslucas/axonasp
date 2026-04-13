# AddFolder Method

## Overview
Adds a local folder hierarchy to the newly created G3TAR archive.

## Syntax
```asp
Dim success
success = obj.AddFolder(folderPath, archiveRoot)
```

## Parameters and Arguments
- folderPath (String, Required): The absolute or relative local path to the directory on the server disk.
- archiveRoot (String, Optional): An alternative folder name under which these files will reside inside the TAR archive.

## Return Values
Returns a `Boolean` representing the operation outcome. Returns True if the folder contents were appended successfully, otherwise False.

## Remarks
- Must be called after initializing an archive using the Create method.
- Calling this method automatically cascades into all subdirectories.

## Code Example
```asp
<%
Option Explicit
Dim obj, success
Set obj = Server.CreateObject("G3TAR")
If obj.Create("C:\temp\backup.tar") Then
    success = obj.AddFolder("C:\temp\reports", "reports_backup")
    If success Then
        Response.Write "Folder added!"
    End If
    obj.Close()
End If
Set obj = Nothing
%>
```