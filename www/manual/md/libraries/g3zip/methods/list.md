# List Method

## Overview
Returns a collection of all file and directory names present in the active archive in the G3Pix AxonASP G3ZIP library.

## Syntax
```asp
fileNamesArray = zip.List()
```

## Return Values
Returns a **VBArray** (Variant) containing the names of all entries in the archive. If the archive is empty or the object is not in Read mode, it returns an empty array.

## Remarks
- The object must be in **Read** mode.
- Directory entries are typically returned with a trailing slash (e.g., "images/").

## Code Example
```asp
<%
Dim zip, files, i
Set zip = Server.CreateObject("G3ZIP")
If zip.Open("/uploads/data.zip") Then
    files = zip.List()
    For i = 0 To UBound(files)
        Response.Write "Entry: " & files(i) & "<br>"
    Next
    zip.Close
End If
Set zip = Nothing
%>
```
