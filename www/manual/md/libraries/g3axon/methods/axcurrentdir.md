# axcurrentdir

## Overview
Retrieves the absolute path of the current working directory for the G3Pix AxonASP process.

## Syntax
```asp
result = obj.axcurrentdir()
```

## Parameters and Arguments
None.

## Return Values
Returns a String containing the absolute path of the current working directory.

## Remarks
If an error occurs while retrieving the directory, an empty string is returned.

## Code Example
```asp
<%
Option Explicit
Dim obj, currentDir
Set obj = Server.CreateObject("G3AXON.FUNCTIONS")

currentDir = obj.axcurrentdir()
Response.Write "The current working directory is: " & currentDir

Set obj = Nothing
%>
```
