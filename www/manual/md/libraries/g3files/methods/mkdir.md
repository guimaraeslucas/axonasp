# Mkdir Method

## Overview
Creates a new directory or a full directory path recursively.

## Syntax
```asp
success = files.Mkdir(path)
```

## Parameters and Arguments
- **path** (String, Required): The directory path to be created.

## Return Values
Returns a **Boolean** indicating whether the directory creation was successful.

## Remarks
- This method is recursive; it will create any missing parent directories in the path automatically.
- If the directory already exists, the method returns **True**.
- This method is also accessible via the **MakeDir** alias.

## Code Example
```asp
<%
Dim files
Set files = Server.CreateObject("G3FILES")
' Creates the entire directory structure recursively
If files.Mkdir("/data/logs/archives/2026") Then
    Response.Write "Directory path created."
End If
Set files = Nothing
%>
```
