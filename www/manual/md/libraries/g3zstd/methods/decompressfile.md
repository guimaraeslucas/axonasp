# DecompressFile Method

## Overview

Decompresses a Zstandard (zstd) compressed file and writes the result to a target file. This G3Pix AxonASP method uses streaming to keep memory usage low even for large files.

## Syntax

```asp
result = obj.DecompressFile(sourcePath, targetPath)
```

## Parameters and Arguments

- **sourcePath**: A String specifying the path to the compressed (.zst) file.
- **targetPath**: A String specifying the path where the decompressed file will be created.

## Return Values

Returns a Boolean value:
- **True**: The file was successfully decompressed.
- **False**: An error occurred during decompression (e.g., file not found, corrupt data).

## Remarks

- Both source and target paths are resolved relative to the web root and must stay within the sandbox.
- This method automatically creates the target directory if it does not exist.
- Streaming ensures that memory consumption is minimal regardless of the file size.
- Errors are logged to the `LastError` property and raise runtime exceptions.

## Code Example

```asp
<%
Option Explicit
Dim objZstd
Set objZstd = Server.CreateObject("G3ZSTD")

' Decompress a file
If objZstd.DecompressFile("data/log.txt.zst", "data/log_restored.txt") Then
    Response.Write "File decompressed successfully."
Else
    Response.Write "Error: " & objZstd.LastError
End If

Set objZstd = Nothing
%>
```
