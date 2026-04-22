# ExtractSingle Method

## Overview
Extract one specific file entry from a G3FC archive.

## Syntax

```asp
result = fc.ExtractSingle(archivePath, entryPath, outputFilePath [, password])
```

## Parameters and Arguments

- archivePath (String, Required): Source `.g3fc` file path.
- entryPath (String, Required): Internal archive entry path to extract.
- outputFilePath (String, Required): Destination file path for the extracted entry.
- password (String, Optional): Archive password for encrypted archives.

## Return Values

- Returns `True` when the target entry is extracted successfully.
- Returns `False` when required arguments are missing, path resolution fails, the entry does not exist, or extraction fails.

## Remarks

- Method names are case-insensitive.
- This method also accepts `extract-single` and `extract_single` aliases.
- Runtime extraction failures raise an internal VBScript error and the method still returns `False`.

## Code Example

```asp
<%
Option Explicit
Dim fc, ok
Set fc = Server.CreateObject("G3FC")

ok = fc.ExtractSingle("/sandbox/archive.g3fc", "docs/readme.txt", "/sandbox/readme.txt", "AxonPass")

If ok Then
    Response.Write "Single entry extracted successfully."
Else
    Response.Write "Single entry extraction failed."
End If

Set fc = Nothing
%>
```





