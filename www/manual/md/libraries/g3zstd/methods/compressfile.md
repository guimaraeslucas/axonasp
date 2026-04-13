# CompressFile Method

## Overview

Compresses a source file and writes the result to a target file using the Zstandard (zstd) algorithm. This G3Pix AxonASP method uses streaming to keep memory usage low even for large files.

## Syntax

```asp
result = obj.CompressFile(sourcePath, targetPath [, level])
```

## Parameters and Arguments

- **sourcePath**: A String specifying the path to the source file.
- **targetPath**: A String specifying the path where the compressed file will be created.
- **level** (Optional): An Integer specifying the compression level. The range is -5 to 22. If omitted, the default level is used.

## Return Values

Returns a Boolean value:
- **True**: The file was successfully compressed.
- **False**: An error occurred during compression (e.g., file not found, permission denied).

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

' Compress a file with level 9
If objZstd.CompressFile("data/log.txt", "data/log.txt.zst", 9) Then
    Response.Write "File compressed successfully."
Else
    Response.Write "Error: " & objZstd.LastError
End If

Set objZstd = Nothing
%>
```
