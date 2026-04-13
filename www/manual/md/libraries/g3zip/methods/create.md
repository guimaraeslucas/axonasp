# Create Method

## Overview
Initializes a new ZIP archive file on the server and prepares the G3Pix AxonASP G3ZIP library for writing operations.

## Syntax
```asp
success = zip.Create(archivePath)
```

## Parameters and Arguments
- **archivePath** (String, Required): The target path where the ZIP file will be created.

## Return Values
Returns a **Boolean** indicating whether the file was successfully created and initialized for writing.

## Remarks
- If the target directory does not exist, the library attempts to create it recursively.
- If a file already exists at the specified path, it will be overwritten.
- Calling **Create** will close any archive currently managed by the object.

## Code Example
```asp
<%
Dim zip
Set zip = Server.CreateObject("G3ZIP")
If zip.Create("/data/exports/data.zip") Then
    Response.Write "Write-mode initialized."
    ' Add files here...
    zip.Close
End If
Set zip = Nothing
%>
```
