# Decompress a File Using G3ZLIB

## Overview

Decompresses a single compressed file back into its original content.

## Syntax

```asp
Dim success
success = obj.DecompressFile(inputPath, outputPath)
```

## Parameters and Arguments

- inputPath (String, Required): The full path to the compressed source file.
- outputPath (String, Required): The destination path for the decompressed file.

## Return Values

Returns a Boolean value. Returns True if the file was decompressed and saved successfully; otherwise, returns False and updates the LastError property.

## Remarks

- Method names are case-insensitive.
- Ensure the destination path allows write operations.

## Code Example

```asp
<%
Option Explicit
Dim obj, success
Set obj = Server.CreateObject("G3ZLIB")

success = obj.DecompressFile(Server.MapPath("document.zlib"), Server.MapPath("extracted_document.txt"))

If success Then
    Response.Write "File decompressed successfully."
Else
    Response.Write "Error: " & obj.LastError
End If

Set obj = Nothing
%>
```



