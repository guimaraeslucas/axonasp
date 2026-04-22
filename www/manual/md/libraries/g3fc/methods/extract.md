# Extract Method

## Overview
Extract all entries from a G3FC archive into a target directory.

## Syntax

```asp
result = fc.Extract(archivePath, outputFolder [, password])
```

## Parameters and Arguments

- archivePath (String, Required): Source `.g3fc` file path.
- outputFolder (String, Required): Destination directory path.
- password (String, Optional): Archive password used when the archive is encrypted.

## Return Values

- Returns `True` when extraction completes successfully.
- Returns `False` when arguments are missing, path resolution fails, or extraction fails.

## Remarks

- Method names are case-insensitive.
- Runtime extraction failures raise an internal VBScript error and the method still returns `False`.

## Code Example

```asp
<%
Option Explicit
Dim fc, ok
Set fc = Server.CreateObject("G3FC")

ok = fc.Extract("/sandbox/archive.g3fc", "/sandbox/extracted", "AxonPass")

If ok Then
    Response.Write "Archive extracted successfully."
Else
    Response.Write "Archive extraction failed."
End If

Set fc = Nothing
%>
```





