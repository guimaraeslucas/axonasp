# Compress a File Using G3ZLIB

## Overview

Compresses a single source file into an output file.

## Syntax

```asp
Dim success
success = obj.CompressFile(inputPath, outputPath, level)
```

## Parameters and Arguments

- inputPath (String, Required): The full path to the source file to compress.
- outputPath (String, Required): The destination path for the compressed file.
- level (Integer, Optional): The compression level from 1 (fastest) to 9 (best compression).

## Return Values

Returns a Boolean value. Returns True if the file was compressed and saved successfully; otherwise, returns False and updates the LastError property.

## Remarks

- Method names are case-insensitive.
- The web server must have read permissions for the input file and write permissions for the output path.

## Code Example

```asp
<%
Option Explicit
Dim obj, success
Set obj = Server.CreateObject("G3ZLIB")

success = obj.CompressFile(Server.MapPath("document.txt"), Server.MapPath("document.zlib"), 9)

If success Then
    Response.Write "File compressed successfully."
Else
    Response.Write "Error: " & obj.LastError
End If

Set obj = Nothing
%>
```



