# Find Method

## Overview
Search archive entries by substring or regular expression.

## Syntax

```asp
matches = fc.Find(archivePath, pattern [, password] [, useRegex])
```

## Parameters and Arguments

- archivePath (String, Required): Source `.g3fc` file path.
- pattern (String, Required): Match expression applied to entry paths.
- password (String, Optional): Archive password for encrypted archives.
- useRegex (Boolean, Optional):
  - `False` (default): uses case-insensitive substring matching.
  - `True`: attempts case-insensitive regular expression matching.

## Return Values

- Returns a zero-based Array of `Scripting.Dictionary` objects on success.
- Each Dictionary includes exactly two keys: `Path` (String) and `Size` (Integer).
- Returns an empty Array when no entries match.
- Returns `Empty` when required arguments are missing, path resolution fails, or archive index reading fails.

## Remarks

- Method names are case-insensitive.
- Regex matching is case-insensitive.
- If `useRegex=True` and the expression fails to compile, matching falls back to case-insensitive substring behavior.
- Runtime read failures raise an internal VBScript error and the method returns `Empty`.

## Code Example

```asp
<%
Option Explicit
Dim fc, matches, i, item
Set fc = Server.CreateObject("G3FC")

matches = fc.Find("/sandbox/archive.g3fc", "\\.txt$", "AxonPass", True)

If IsArray(matches) Then
    For i = LBound(matches) To UBound(matches)
        Set item = matches(i)
        Response.Write item.Item("Path") & "<br>"
        Set item = Nothing
    Next
Else
    Response.Write "Find failed."
End If

Set fc = Nothing
%>
```





