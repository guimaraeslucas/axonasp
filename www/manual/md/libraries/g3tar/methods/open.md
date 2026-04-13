# Open Method

## Overview
Accesses an established TAR archive on the server to read or extract its contents.

## Syntax
```asp
Dim success
success = obj.Open(archivePath)
```

## Parameters and Arguments
- archivePath (String, Required): Exact system location indicating the existing TAR file on the filesystem.

## Return Values
Returns a `Boolean` representing whether the operation locked the file properly. Returns True if successfully mounted.

## Remarks
- Must be combined with a paired extraction routine or metadata gathering phase followed immediately by a manual invocation of Close.
- Do not use this method on archives concurrently accessed through other means on the local disk.

## Code Example
```asp
<%
Option Explicit
Dim obj, success
Set obj = Server.CreateObject("G3TAR")
success = obj.Open("C:\data\packages\archive.tar")
If success Then
    Response.Write "Successfully connected."
    obj.Close()
End If
Set obj = Nothing
%>
```