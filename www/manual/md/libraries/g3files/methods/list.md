# List Method

## Overview
Returns an Array containing the names of all files within a specified directory.

## Syntax
```asp
fileNamesArray = files.List(path)
```

## Parameters and Arguments
- **path** (String, Required): The target directory path to list.

## Return Values
Returns an **Array** (Variant) containing the names of the files. If the directory is empty or does not exist, it returns an empty array.

## Remarks
- The returned array only includes file names, not subdirectories.
- Path resolution is relative to the AxonASP sandbox root.
- This method is also accessible via the **ListFiles** alias.

## Code Example
```asp
<%
Dim files, fileList, i
Set files = Server.CreateObject("G3FILES")
fileList = files.List("/uploads")

If UBound(fileList) >= 0 Then
    Response.Write "Found " & UBound(fileList) + 1 & " files:<br>"
    For i = 0 To UBound(fileList)
        Response.Write "- " & fileList(i) & "<br>"
    Next
Else
    Response.Write "No files found in directory."
End If
Set files = Nothing
%>
```
