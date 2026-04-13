# GetInfo Method

## Overview
Compiles and maps extensive metadata fields referring to an individual file path stored within the presently loaded TAR structure.

## Syntax
```asp
Dim dict
Set dict = obj.GetInfo(archiveName)
```

## Parameters and Arguments
- archiveName (String, Required): Target relative naming referring to the file entry inside the TAR file.

## Return Values
Yields a `Scripting.Dictionary` native interface holding extracted fields. Should an anomaly manifest or if no such element occurs, an instantiated blank model will emerge.

## Remarks
- Necessitates pre-opening via the Open invocation.
- Valuable to verify sizing specifications alongside creation timing metadata properties inside the internal container sequence before actual unpacking logic begins.

## Code Example
```asp
<%
Option Explicit
Dim obj, fileStats
Set obj = Server.CreateObject("G3TAR")
If obj.Open("C:\data\packages\archive.tar") Then
    Set fileStats = obj.GetInfo("docs/readme.txt")
    If fileStats.Count > 0 Then
        Response.Write "Identified item data bounds correctly."
    End If
    Set fileStats = Nothing
    obj.Close()
End If
Set obj = Nothing
%>
```