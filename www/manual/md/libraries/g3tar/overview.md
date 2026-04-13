# Use G3TAR in AxonASP

## Overview
G3TAR provides robust TAR archive creation and extraction capabilities. It allows packaging files and folders into an uncompressed tape archive format or extracting existing archives to a local directory.

## Syntax
```asp
Dim obj
Set obj = Server.CreateObject("G3TAR")
```

## Parameters and Arguments
- ProgID (String, Required): Must initialize with "G3TAR".

## Return Values
Returns a native object handle for the G3Pix AxonASP G3TAR engine.

## Remarks
- G3TAR works primarily with the uncompressed standard TAR format. For compressed wrappers, combine with G3ZLIB or G3ZSTD operations if needed.
- Always check the returned boolean values from methods and read the LastError property if an operation fails.
- Do not forget to invoke Close after finishing operations to release file handles.

## Code Example
```asp
<%
Option Explicit
Dim obj
Set obj = Server.CreateObject("G3TAR")
If obj.Create("C:\temp\backup.tar") Then
    obj.AddFile "C:\temp\data.txt", "data.txt"
    obj.Close()
End If
Set obj = Nothing
%>
```