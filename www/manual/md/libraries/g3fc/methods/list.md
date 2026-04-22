# List Method

## Overview
List entries stored in a G3FC archive.

## Syntax

```asp
entries = fc.List(archivePath [, password] [, unit] [, details])
```

## Parameters and Arguments

- archivePath (String, Required): Source `.g3fc` file path.
- password (String, Optional): Archive password for encrypted archives.
- unit (String, Optional): Formatted size unit used by `FormattedSize`. Supported values are `TB`, `GB`, `MB`, and `KB`. Default is `KB`. Any other value returns sizes as `B` text.
- details (Boolean, Optional): When `True`, includes extended metadata fields in each result item. Default is `False`.

## Return Values

- Returns a zero-based Array of `Scripting.Dictionary` objects when listing succeeds.
- Each Dictionary always includes: `Path` (String), `Size` (Integer), `FormattedSize` (String), and `Type` (String).
- When `details=True`, each Dictionary also includes: `Permissions` (String in octal format), `CreationTime` (RFC3339 String), and `Checksum` (uppercase 8-digit hexadecimal String).
- Returns `Empty` when required arguments are missing, path resolution fails, or archive index reading fails.

## Remarks

- Method names are case-insensitive.
- Runtime read failures raise an internal VBScript error and the method returns `Empty`.

## Code Example

```asp
<%
Option Explicit
Dim fc, entries, i, item
Set fc = Server.CreateObject("G3FC")

entries = fc.List("/sandbox/archive.g3fc", "AxonPass", "MB", True)

If IsArray(entries) Then
    For i = LBound(entries) To UBound(entries)
        Set item = entries(i)
        Response.Write item.Item("Path") & " - " & item.Item("FormattedSize") & "<br>"
        Set item = Nothing
    Next
Else
    Response.Write "List failed."
End If

Set fc = Nothing
%>
```





