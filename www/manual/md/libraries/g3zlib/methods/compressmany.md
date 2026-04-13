# Compress Multiple Files Using G3ZLIB

## Overview

Compresses an array or list of source files into a single output bundle file.

## Syntax

```asp
Dim success
success = obj.CompressMany(sourcePaths, outputPath, level)
```

## Parameters and Arguments

- sourcePaths (Variant, Required): An array of strings representing the file paths to compress.
- outputPath (String, Required): The destination path for the compressed bundle file.
- level (Integer, Optional): The compression level from 1 (fastest) to 9 (best compression). Defaults to the standard level if omitted.

## Return Values

Returns a Boolean value. Returns True if the files were compressed and saved successfully; otherwise, returns False and updates the LastError property.

## Remarks

- Method names are case-insensitive.
- Ensure that the web server has the appropriate read permissions for the source paths and write permissions for the output directory.

## Code Example

```asp
<%
Option Explicit
Dim obj, success, files(1)
files(0) = Server.MapPath("file1.txt")
files(1) = Server.MapPath("file2.txt")

Set obj = Server.CreateObject("G3ZLIB")
success = obj.CompressMany(files, Server.MapPath("archive.zlib"), 9)

If success Then
    Response.Write "Files compressed successfully."
Else
    Response.Write "Error: " & obj.LastError
End If

Set obj = Nothing
%>
```



