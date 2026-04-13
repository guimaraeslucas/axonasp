# Close Method

## Overview
Finalizes the archive operations and releases all system resources in the G3Pix AxonASP G3ZIP library.

## Syntax
```asp
zip.Close()
```

## Return Values
Returns a **Boolean** (True) upon completion.

## Remarks
- For **Write** mode, this method is mandatory to ensure all buffers are flushed and the ZIP header is correctly written to disk.
- For **Read** mode, it closes the file handle and releases the archive reader.
- After calling **Close**, the **Mode** and **Path** properties are reset.

## Code Example
```asp
<%
Dim zip
Set zip = Server.CreateObject("G3ZIP")
zip.Create("/temp/output.zip")
' ... perform operations ...
zip.Close
Set zip = Nothing
%>
```
