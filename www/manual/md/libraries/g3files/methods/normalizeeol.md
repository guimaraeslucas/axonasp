# NormalizeEOL Method

## Overview
Returns a string with line endings converted to a specified style.

## Syntax
```asp
normalizedText = files.NormalizeEOL(text, style)
```

## Parameters and Arguments
- **text** (String, Required): The source text containing inconsistent line endings.
- **style** (String, Required): The target line ending style (e.g., "windows", "linux", "mac").

## Return Values
Returns a **String** with the normalized line endings.

## Remarks
- Styles include "windows" (CRLF), "linux" (LF), and "mac" (CR).
- This method is performed in memory and does not affect any files on disk.
- This method is also accessible via the **NormalizeLineEndings** alias.

## Code Example
```asp
<%
Dim files, rawText, normalizedText
Set files = Server.CreateObject("G3FILES")
rawText = "Line 1" & Chr(10) & "Line 2" & Chr(13) & "Line 3"
normalizedText = files.NormalizeEOL(rawText, "windows")
' All line endings are now CRLF
Set files = Nothing
%>
```
