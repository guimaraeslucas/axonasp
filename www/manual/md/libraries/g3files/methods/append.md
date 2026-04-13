# Append Method

## Overview
Appends text content to an existing file or creates a new file if it does not already exist.

## Syntax
```asp
success = files.Append(path, content [, encoding] [, lineEnding])
```

## Parameters and Arguments
- **path** (String, Required): The target file path relative to the AxonASP sandbox.
- **content** (String, Required): The text to append to the file.
- **encoding** (String, Optional): The text encoding to use (e.g., "utf-8", "utf-16", "ascii", "iso-8859-1"). The default is "utf-8".
- **lineEnding** (String, Optional): The line ending style to apply (e.g., "windows", "linux", "mac"). The default is "linux" (LF).

## Return Values
Returns a **Boolean** indicating whether the append operation was successful.

## Remarks
- If the file exists, the new content is added to the end.
- If the file does not exist, it is created.
- The library automatically handles directory creation if the target path includes missing folders.
- This method is also accessible via the **AppendText** alias.

## Code Example
```asp
<%
Dim files, logPath
Set files = Server.CreateObject("G3FILES")
logPath = "/logs/app_log.txt"
If files.Append(logPath, Now() & " - User logged in", "utf-8", "windows") Then
    Response.Write "Log entry added."
End If
Set files = Nothing
%>
```
