# Delete Method

## Overview
Permanently removes a file or directory from the file system.

## Syntax
```asp
success = files.Delete(path)
```

## Parameters and Arguments
- **path** (String, Required): The target file or directory path to be removed.

## Return Values
Returns a **Boolean** indicating whether the deletion was successful.

## Remarks
- If the target path is a directory, it must be empty before it can be deleted.
- This method is also accessible via the **Remove** alias.
- Path resolution is relative to the AxonASP sandbox root.

## Code Example
```asp
<%
Dim files
Set files = Server.CreateObject("G3FILES")
If files.Delete("/temp/old_data.txt") Then
    Response.Write "File deleted."
Else
    Response.Write "Failed to delete file."
End If
Set files = Nothing
%>
```
