# Move Method

## Overview
Renames or moves a file or directory to a new location in the file system.

## Syntax
```asp
success = files.Move(source, dest)
```

## Parameters and Arguments
- **source** (String, Required): The path to the file or directory to be moved.
- **dest** (String, Required): The target path for the move or rename.

## Return Values
Returns a **Boolean** indicating whether the move or rename operation was successful.

## Remarks
- If the destination already exists, the move operation will fail.
- This method is also accessible via the **Rename** alias.
- Path resolution is relative to the AxonASP sandbox root.

## Code Example
```asp
<%
Dim files
Set files = Server.CreateObject("G3FILES")
If files.Move("/temp/data.csv", "/data/processed/data.csv") Then
    Response.Write "File moved."
End If
Set files = Nothing
%>
```
