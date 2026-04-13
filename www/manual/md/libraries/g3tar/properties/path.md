# Path Property

## Overview
Gets the active file path associated with the G3TAR archive.

## Syntax
```asp
Dim filePath
filePath = obj.Path
```

## Parameters and Arguments
- Getter: None.

## Return Values
Returns a `String` containing the absolute path of the current archive.

## Remarks
- This property is read-only.
- Available only after opening or creating an archive.

## Code Example
```asp
<%
Option Explicit
Dim obj, filePath
Set obj = Server.CreateObject("G3TAR")
If obj.Open("C:\temp\archive.tar") Then
    filePath = obj.Path
    Response.Write filePath
End If
obj.Close()
Set obj = Nothing
%>
```