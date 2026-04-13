# AddFile Method

## Overview
Includes a physical file into the current write-mode archive in the G3Pix AxonASP G3ZIP library.

## Syntax
```asp
success = zip.AddFile(sourcePath [, nameInZip])
```

## Parameters and Arguments
- **sourcePath** (String, Required): The path to the file on the server to be added.
- **nameInZip** (String, Optional): The relative path and name the file will have inside the ZIP archive. If omitted, the library uses the base name of the source file.

## Return Values
Returns a **Boolean** indicating whether the file was successfully added to the archive.

## Remarks
- The object must be in **Write** mode (initialized via the **Create** method).
- Leading slashes in `nameInZip` are automatically removed.
- Path resolution for `sourcePath` is handled relative to the AxonASP sandbox.

## Code Example
```asp
<%
Dim zip
Set zip = Server.CreateObject("G3ZIP")
If zip.Create("/temp/archive.zip") Then
    ' Add a file and rename it within the ZIP
    zip.AddFile "/images/logo.png", "assets/branding.png"
    zip.Close
End If
Set zip = Nothing
%>
```
