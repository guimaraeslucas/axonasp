# Write Method

## Overview
Writes text content to a file, overwriting any existing content.

## Syntax
```asp
success = files.Write(path, content [, encoding] [, lineEnding] [, includeBOM])
```

## Parameters and Arguments
- **path** (String, Required): The target file path.
- **content** (String, Required): The text content to write.
- **encoding** (String, Optional): The text encoding to use (e.g., "utf-8", "utf-16", "ascii", "iso-8859-1"). The default is "utf-8".
- **lineEnding** (String, Optional): The line ending style to apply (e.g., "windows", "linux"). The default is "linux" (LF).
- **includeBOM** (Boolean, Optional): Specifies whether to include a Byte Order Mark in the file. The default is **False**.

## Return Values
Returns a **Boolean** indicating whether the write operation was successful.

## Remarks
- If the file already exists, it is overwritten.
- The library automatically handles directory creation if the target path includes missing folders.
- This method is also accessible via the **WriteText** alias.

## Code Example
```asp
<%
Dim files, content, success
Set files = Server.CreateObject("G3FILES")
content = "Configuration Data" & vbCrLf & "Version: 2.1"
' Write as UTF-8 with a BOM and Windows line endings
success = files.Write("config.txt", content, "utf-8", "windows", True)
If success Then
    Response.Write "Configuration saved."
End If
Set files = Nothing
%>
```
