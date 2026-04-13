# AddFiles Method

## Overview
Appends a list of specific files to the current G3TAR archive.

## Syntax
```asp
Dim success
success = obj.AddFiles(sourcePaths, prefix)
```

## Parameters and Arguments
- sourcePaths (Variant, Required): An iterable, such as an Array or a Scripting.Dictionary, containing the paths to append.
- prefix (String, Optional): A common folder path to prepend to the archived names for grouping.

## Return Values
Returns a `Boolean` indicating if all files in the provided structure were successfully appended. True on total success.

## Remarks
- Must be preceded by a successful call to the Create method.
- Useful when selectively archiving files without needing to move them into a consolidated directory first on the server disk.

## Code Example
```asp
<%
Option Explicit
Dim obj, success, paths(1)
paths(0) = "C:\temp\summary.txt"
paths(1) = "C:\temp\metrics.csv"

Set obj = Server.CreateObject("G3TAR")
If obj.Create("C:\temp\combined.tar") Then
    success = obj.AddFiles(paths, "dataset")
    If success Then
        Response.Write "Items added!"
    End If
    obj.Close()
End If
Set obj = Nothing
%>
```