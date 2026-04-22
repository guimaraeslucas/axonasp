# Create Method

## Overview
Create a new G3FC archive from one file path or multiple source paths.

## Syntax

```asp
result = fc.Create(archivePath, sourcePaths [, password] [, options])
```

## Parameters and Arguments

- archivePath (String, Required): Output `.g3fc` file path. The path is resolved through the current ASP host root.
- sourcePaths (String or Array, Required): Source input. Pass one path as String, or pass a VBScript Array of paths.
- password (String, Optional): Read password written into the archive configuration. When non-empty, encryption mode is enabled.
- options (Scripting.Dictionary, Optional): Configuration Dictionary. Supported keys:
  - `CompressionLevel` (Integer): Zstandard compression level. Default is `6`.
  - `GlobalCompression` (Boolean): If `True`, keeps file blocks uncompressed and uses global compression mode.
  - `FECLevel` (Integer): Forward error correction level. Values greater than `0` enable FEC.
  - `SplitSize` (String): Segment size in `MB` or `GB` format (example: `100MB`, `1GB`).

## Return Values

- Returns `True` when archive creation completes successfully.
- Returns `False` when any required argument is missing, output path resolution fails, no valid source files are resolved, or archive creation fails.

## Remarks

- Method names are case-insensitive.
- Invalid source paths are skipped during source resolution.
- If no valid source file remains after resolution, the method returns `False`.
- Runtime failures raise an internal VBScript error and the method still returns `False`.

## Code Example

```asp
<%
Option Explicit
Dim fc, options, ok
Set fc = Server.CreateObject("G3FC")
Set options = Server.CreateObject("Scripting.Dictionary")

Call options.Add("CompressionLevel", 9)
Call options.Add("GlobalCompression", False)
Call options.Add("FECLevel", 10)
Call options.Add("SplitSize", "100MB")

ok = fc.Create("/sandbox/archive.g3fc", Array("/sandbox/a.txt", "/sandbox/b.txt"), "AxonPass", options)

If ok Then
    Response.Write "Archive created successfully."
Else
    Response.Write "Archive creation failed."
End If

Set options = Nothing
Set fc = Nothing
%>
```





