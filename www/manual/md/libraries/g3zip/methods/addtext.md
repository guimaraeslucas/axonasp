# AddText Method

## Overview
Creates a new virtual file inside the current write-mode archive using a provided string in the G3Pix AxonASP G3ZIP library.

## Syntax
```asp
success = zip.AddText(nameInZip, content)
```

## Parameters and Arguments
- **nameInZip** (String, Required): The relative path and name for the new file inside the archive.
- **content** (String, Required): The text content to be written into the file.

## Return Values
Returns a **Boolean** indicating whether the virtual file was successfully created and added.

## Remarks
- This method is ideal for generating dynamic files such as manifest.json, readme.txt, or configuration files on the fly.
- The object must be in **Write** mode.

## Code Example
```asp
<%
Dim zip, info
Set zip = Server.CreateObject("G3ZIP")
If zip.Create("/temp/report.zip") Then
    info = "Report generated on: " & Now() & vbCrLf & "System: AxonASP"
    zip.AddText "metadata.txt", info
    zip.Close
End If
Set zip = Nothing
%>
```
