# Count Property

## Overview
Returns the total number of files and directories currently contained in the active ZIP archive in the G3Pix AxonASP G3ZIP library.

## Syntax
```asp
itemCount = zip.Count
```

## Return Values
Returns an **Integer** representing the entry count. If no archive is opened in Read mode, it returns 0.

## Remarks
- This property is only populated when the object is in **Read** mode (initialized via the **Open** method).
- It is read-only.

## Code Example
```asp
<%
Dim zip
Set zip = Server.CreateObject("G3ZIP")
If zip.Open("/assets/data.zip") Then
    Response.Write "Archive entries: " & zip.Count
    zip.Close
End If
Set zip = Nothing
%>
```
