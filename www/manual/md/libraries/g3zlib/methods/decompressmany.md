# Decompress Multiple Files Using G3ZLIB

## Overview

Decompresses a bundled ZLIB archive back into multiple files in a specified destination directory.

## Syntax

```asp
Dim success
success = obj.DecompressMany(inputPath, outputFolder)
```

## Parameters and Arguments

- inputPath (String, Required): The path to the compressed bundle file.
- outputFolder (String, Required): The directory where the decompressed files will be saved.

## Return Values

Returns a Boolean value. Returns True if all files were decompressed and saved successfully; otherwise, returns False and updates the LastError property.

## Remarks

- Method names are case-insensitive.
- Ensure that the output folder exists and that the web server has write permissions to it.

## Code Example

```asp
<%
Option Explicit
Dim obj, success
Set obj = Server.CreateObject("G3ZLIB")

success = obj.DecompressMany(Server.MapPath("archive.zlib"), Server.MapPath("output_folder\"))

If success Then
    Response.Write "Archive decompressed successfully."
Else
    Response.Write "Error: " & obj.LastError
End If

Set obj = Nothing
%>
```



