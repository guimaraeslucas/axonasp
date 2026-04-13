# Open Method

## Overview
Opens an existing ZIP archive file for reading and inspection in the G3Pix AxonASP G3ZIP library.

## Syntax
```asp
success = zip.Open(archivePath)
```

## Parameters and Arguments
- **archivePath** (String, Required): The path to the ZIP file on the server.

## Return Values
Returns a **Boolean** indicating whether the archive was successfully opened.

## Remarks
- Calling **Open** will close any archive currently managed by the object.
- Once opened, the object enters **Read** mode, enabling methods like **List**, **GetInfo**, and **ExtractAll**.

## Code Example
```asp
<%
Dim zip
Set zip = Server.CreateObject("G3ZIP")
If zip.Open("/data/source.zip") Then
    Response.Write "Archive has " & zip.Count & " files."
    zip.Close
End If
Set zip = Nothing
%>
```
