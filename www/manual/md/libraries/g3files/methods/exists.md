# Exists Method

## Overview
Returns a Boolean indicating whether a file or directory exists at the specified path.

## Syntax
```asp
boolExists = files.Exists(path)
```

## Parameters and Arguments
- **path** (String, Required): The target path to check.

## Return Values
Returns a **Boolean** value. It returns **True** if the file or directory exists, and **False** otherwise.

## Remarks
- Path resolution is relative to the AxonASP sandbox root.
- This method is efficient for validating resources before attempting read or write operations.

## Code Example
```asp
<%
Dim files
Set files = Server.CreateObject("G3FILES")
If files.Exists("/web.config") Then
    Response.Write "Configuration file exists."
Else
    Response.Write "File not found."
End If
Set files = Nothing
%>
```
