# Count Property

## Overview
Gets the total number of entries within the active G3TAR archive.

## Syntax
```asp
Dim entryCount
entryCount = obj.Count
```

## Parameters and Arguments
- Getter: None.

## Return Values
Returns an `Integer` representing the total files and folders tracked.

## Remarks
- This property is read-only.
- Returns zero if no archive is open.

## Code Example
```asp
<%
Option Explicit
Dim obj, entryCount
Set obj = Server.CreateObject("G3TAR")
If obj.Open("C:\temp\archive.tar") Then
    entryCount = obj.Count
    Response.Write entryCount
End If
obj.Close()
Set obj = Nothing
%>
```

