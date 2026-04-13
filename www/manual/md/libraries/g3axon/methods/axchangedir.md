# axchangedir

## Overview
Changes the current working directory of the process in G3Pix AxonASP.

## Syntax
```asp
result = obj.axchangedir(path)
```

## Parameters and Arguments
- **path** (String): The absolute or relative path to the new working directory.

## Return Values
Returns a Boolean indicating whether the directory change was successful.

## Remarks
Changing the working directory affects the entire process. In a web server environment, this should be used with caution as it may affect other concurrent requests.

## Code Example
```asp
<%
Option Explicit
Dim obj, result
Set obj = Server.CreateObject("G3AXON.FUNCTIONS")

result = obj.axchangedir("C:\Windows\Temp")

If result Then
    Response.Write "Successfully changed directory to: " & obj.axcurrentdir()
Else
    Response.Write "Failed to change directory."
End If

Set obj = Nothing
%>
```
