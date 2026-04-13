# Read Method

## Overview
Returns the full text content of a file using a specified encoding.

## Syntax
```asp
fileContent = files.Read(path [, encoding])
```

## Parameters and Arguments
- **path** (String, Required): The path to the file to be read.
- **encoding** (String, Optional): The text encoding to use for decoding the file (e.g., "utf-8", "utf-16", "ascii", "iso-8859-1"). The default is "utf-8".

## Return Values
Returns a **String** containing the full content of the file. If the file cannot be read, it returns an empty string.

## Remarks
- The library automatically detects Byte Order Marks (BOM) to determine the correct encoding.
- This method is also accessible via the **ReadText** alias.
- Path resolution is relative to the AxonASP sandbox root.

## Code Example
```asp
<%
Dim files, content
Set files = Server.CreateObject("G3FILES")
content = files.Read("/config/app.json", "utf-8")
If content <> "" Then
    Response.Write "File read successfully."
End If
Set files = Nothing
%>
```
