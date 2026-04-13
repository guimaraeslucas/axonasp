# ExtractAll Method

## Overview
Reconstructs the complete TAR folder layout natively across a defined destination volume.

## Syntax
```asp
Dim success
success = obj.ExtractAll(outputFolder)
```

## Parameters and Arguments
- outputFolder (String, Required): Absolute server directory targeting where the complete payload extraction sequence lands.

## Return Values
Returns a `Boolean` equating True if the overall sequence finishes processing without causing integrity blockades or path exceptions.

## Remarks
- Needs a verified session constructed previously targeting the appropriate file handle via Open.
- Automatically handles file access creation permissions when reproducing paths originally saved.

## Code Example
```asp
<%
Option Explicit
Dim obj, success
Set obj = Server.CreateObject("G3TAR")
If obj.Open("C:\temp\data.tar") Then
    success = obj.ExtractAll("C:\unpacked_volumes")
    If success Then
        Response.Write "Sequence finished."
    End If
    obj.Close()
End If
Set obj = Nothing
%>
```