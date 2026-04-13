# ExtractFile Method

## Overview
Decompresses and pushes one designated file from the connected TAR compilation directly onto the host storage medium.

## Syntax
```asp
Dim success
success = obj.ExtractFile(archiveName, outputPath)
```

## Parameters and Arguments
- archiveName (String, Required): Target path naming referring directly to the stored inner artifact.
- outputPath (String, Required): Exact endpoint pointer declaring where the output payload becomes solidified onto disk storage.

## Return Values
Dispatches a `Boolean` expression rendering True for success and False upon error constraints.

## Remarks
- This approach assumes a correct reference initialization on Open usage.
- If folders matching the destination parameters cannot be identified, this sequence will automatically scaffold paths as configured internally.

## Code Example
```asp
<%
Option Explicit
Dim obj, success
Set obj = Server.CreateObject("G3TAR")
If obj.Open("C:\data\packages\files.tar") Then
    success = obj.ExtractFile("source/test.jpg", "C:\data\extracted\test.jpg")
    If success Then
        Response.Write "Extract sequence finalized smoothly."
    End If
    obj.Close()
End If
Set obj = Nothing
%>
```