# GetInfo Method

## Overview
Retrieves detailed technical metadata for a specific file within the archive in the G3Pix AxonASP G3ZIP library.

## Syntax
```asp
Set info = zip.GetInfo(fileName)
```

## Parameters and Arguments
- **fileName** (String, Required): The name of the file inside the ZIP archive.

## Return Values
Returns a **Scripting.Dictionary** object containing the following keys:
- **Name** (String): The name of the file.
- **Size** (Integer): The uncompressed size in bytes.
- **PackedSize** (Integer): The compressed size in bytes.
- **Modified** (String): The last modification timestamp in RFC3339 format.
- **IsDir** (Boolean): Indicates if the entry is a directory.

Returns **Empty** if the file is not found or the object is not in Read mode.

## Remarks
- The object must be in **Read** mode.
- The `fileName` matching is case-insensitive.

## Code Example
```asp
<%
Dim zip, info
Set zip = Server.CreateObject("G3ZIP")
If zip.Open("/temp/data.zip") Then
    Set info = zip.GetInfo("report.csv")
    If Not IsEmpty(info) Then
        Response.Write "File: " & info("Name") & "<br>"
        Response.Write "Uncompressed Size: " & info("Size") & " bytes"
    End If
    zip.Close
End If
Set zip = Nothing
%>
```
