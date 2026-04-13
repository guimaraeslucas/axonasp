# ExtractAll Method

## Overview
Unpacks the entire content of the active read-mode archive into a specified directory in the G3Pix AxonASP G3ZIP library.

## Syntax
```asp
success = zip.ExtractAll(targetDirectory)
```

## Parameters and Arguments
- **targetDirectory** (String, Required): The destination directory on the server where the files will be extracted.

## Return Values
Returns a **Boolean** indicating whether all files were successfully extracted.

## Remarks
- The object must be in **Read** mode (initialized via the **Open** method).
- The library automatically recreates any subdirectory structure found within the archive.
- Existing files in the destination directory will be overwritten if they have the same name.

## Code Example
```asp
<%
Dim zip
Set zip = Server.CreateObject("G3ZIP")
If zip.Open("/uploads/archive.zip") Then
    If zip.ExtractAll("/temp/extracted_files") Then
        Response.Write "Extraction complete."
    End If
    zip.Close
End If
Set zip = Nothing
%>
```
