# Create Method

## Overview
Generates the base file pointer tracking an entirely novel TAR archive mapped structurally onto the server system.

## Syntax
```asp
Dim success
success = obj.Create(archivePath)
```

## Parameters and Arguments
- archivePath (String, Required): Designated absolute path modeling the storage placement expected upon sequence closure.

## Return Values
Submits a `Boolean` representation validating disk write accessibility and the file sequence lock instantiation correctly. Returns True on success, otherwise it fails yielding False.

## Remarks
- Will aggressively truncate or replace similar file identities encountered explicitly named within the given string array coordinates.
- Expects subsequent interactions populating context data leveraging add components beforehand routing back to Close.

## Code Example
```asp
<%
Option Explicit
Dim obj, success
Set obj = Server.CreateObject("G3TAR")
success = obj.Create("C:\temp\new_output.tar")
If success Then
    Response.Write "Initiation created."
    obj.Close()
End If
Set obj = Nothing
%>
```