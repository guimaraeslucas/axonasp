# ConvertFileEncoding Method

## Overview
Converts the text encoding and line endings of a file on disk and saves the result to a destination path.

## Syntax
```asp
success = files.ConvertFileEncoding(source, dest, srcEnc, dstEnc [, lineEnding] [, includeBOM])
```

## Parameters and Arguments
- **source** (String, Required): The path to the source file.
- **dest** (String, Required): The path to the destination file.
- **srcEnc** (String, Required): The encoding of the source file.
- **dstEnc** (String, Required): The target encoding for the destination file.
- **lineEnding** (String, Optional): The target line ending style (e.g., "windows", "linux").
- **includeBOM** (Boolean, Optional): Specifies whether to include a Byte Order Mark in the destination file.

## Return Values
Returns a **Boolean** indicating whether the conversion and save operation were successful.

## Remarks
- This method is useful for batch-processing files for cross-platform compatibility.
- Supported encodings include "utf-8", "utf-16", "ascii", and "iso-8859-1".

## Code Example
```asp
<%
Dim files
Set files = Server.CreateObject("G3FILES")
' Convert an ISO-8859-1 file to UTF-8 with Windows line endings and a BOM
If files.ConvertFileEncoding("legacy.txt", "modern.txt", "iso-8859-1", "utf-8", "windows", True) Then
    Response.Write "File converted successfully."
End If
Set files = Nothing
%>
```
