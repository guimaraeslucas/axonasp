# AddFile Method

## Overview
Adds a single local file to the active G3TAR archive.

## Syntax
```asp
Dim success
success = obj.AddFile(filePath, archiveName)
```

## Parameters and Arguments
- filePath (String, Required): Absolute or relative server-side path pointing to the original file to read.
- archiveName (String, Optional): Target filename or internal directory structure under which the file resides within the TAR archive. If omitted, the base filename is used.

## Return Values
Returns a `Boolean`. Computes to True if the file bytes were flushed to the archive, otherwise False.

## Remarks
- Requires a previous call to the Create method.
- Useful for granular file consolidation across disparate source directories.

## Code Example
```asp
<%
Option Explicit
Dim obj, success
Set obj = Server.CreateObject("G3TAR")
If obj.Create("C:\temp\output.tar") Then
    success = obj.AddFile("C:\logs\system.log", "logs/system.log")
    If success Then
        Response.Write "Item appended properly."
    End If
    obj.Close()
End If
Set obj = Nothing
%>
```