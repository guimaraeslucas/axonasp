# Info Method

## Overview
Export archive metadata from a G3FC file into an output file.

## Syntax

```asp
result = fc.Info(archivePath, outputFilePath [, password])
```

## Parameters and Arguments

- archivePath (String, Required): Source `.g3fc` file path.
- outputFilePath (String, Required): Destination file path that receives the exported metadata.
- password (String, Optional): Archive password for encrypted archives.

## Return Values

- Returns `True` when metadata export succeeds.
- Returns `False` when required arguments are missing, path resolution fails, or export fails.

## Remarks

- Method names are case-insensitive.
- Runtime export failures raise an internal VBScript error and the method still returns `False`.

## Code Example

```asp
<%
Option Explicit
Dim fc, ok
Set fc = Server.CreateObject("G3FC")

ok = fc.Info("/sandbox/archive.g3fc", "/sandbox/archive-info.txt", "AxonPass")

If ok Then
    Response.Write "Archive info exported."
Else
    Response.Write "Archive info export failed."
End If

Set fc = Nothing
%>
```





