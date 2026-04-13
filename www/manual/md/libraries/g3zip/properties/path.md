# Path Property

## Overview
Returns the absolute physical path of the archive file currently managed by the G3Pix AxonASP G3ZIP library.

## Syntax
```asp
fullPath = zip.Path
```

## Return Values
Returns a **String** containing the full path to the ZIP file. Returns an empty string if no archive is active.

## Remarks
- This property is read-only.
- It reflects the resolved path after `Server.MapPath` or relative resolution has been applied by the engine.

## Code Example
```asp
<%
Dim zip
Set zip = Server.CreateObject("G3ZIP")
zip.Open "data.zip"
Response.Write "Resolved path: " & zip.Path
zip.Close
Set zip = Nothing
%>
```
