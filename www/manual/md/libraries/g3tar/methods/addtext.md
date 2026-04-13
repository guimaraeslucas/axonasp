# AddText Method

## Overview
Adds string content as an inline file into the active G3TAR archive.

## Syntax
```asp
Dim success
success = obj.AddText(archiveName, text)
```

## Parameters and Arguments
- archiveName (String, Required): Target relative filename where the text will be stored within the TAR.
- text (String, Required): The string to embed as the file body.

## Return Values
Returns a `Boolean` which resolves to True if the text length was fully preserved and saved correctly.

## Remarks
- Requires a current active archive opened via the Create method.
- Allows immediate injection of dynamically generated metadata, manifests, or JSON reports without dumping to the disk first.

## Code Example
```asp
<%
Option Explicit
Dim obj, success, body
body = "This is a dynamically created file."

Set obj = Server.CreateObject("G3TAR")
If obj.Create("C:\temp\report.tar") Then
    success = obj.AddText("readme.txt", body)
    If success Then
        Response.Write "Text generated straight into archive."
    End If
    obj.Close()
End If
Set obj = Nothing
%>
```