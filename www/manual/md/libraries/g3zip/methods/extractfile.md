# ExtractFile Method

## Overview
Unpacks a specific file from the active read-mode archive into a specified directory in the G3Pix AxonASP G3ZIP library.

## Syntax
```asp
success = zip.ExtractFile(fileName, targetDirectory)
```

## Parameters and Arguments
- **fileName** (String, Required): The name (or relative path) of the file inside the ZIP archive to be extracted.
- **targetDirectory** (String, Required): The destination directory on the server.

## Return Values
Returns a **Boolean** indicating whether the specific file was found and successfully extracted.

## Remarks
- The object must be in **Read** mode.
- The `fileName` parameter is case-sensitive and must match the entry name returned by the **List** method exactly.

## Code Example
```asp
<%
Dim zip
Set zip = Server.CreateObject("G3ZIP")
If zip.Open("/assets/pack.zip") Then
    ' Extract only the documentation
    If zip.ExtractFile("docs/readme.pdf", "/www/help") Then
        Response.Write "Documentation extracted."
    End If
    zip.Close
End If
Set zip = Nothing
%>
```
