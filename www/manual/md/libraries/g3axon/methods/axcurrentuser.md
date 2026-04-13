# axcurrentuser

## Overview
Retrieves the username of the user currently running the G3Pix AxonASP process.

## Syntax
```asp
result = obj.axcurrentuser()
```

## Parameters and Arguments
None.

## Return Values
Returns a String containing the username of the user who owns the current process.

## Remarks
On Windows, it first attempts to retrieve the current user's name using system APIs, falling back to the "USERNAME" environment variable if necessary. On Unix systems, it retrieves the name associated with the process owner.

## Code Example
```asp
<%
Option Explicit
Dim obj, userName
Set obj = Server.CreateObject("G3AXON.FUNCTIONS")

userName = obj.axcurrentuser()
Response.Write "Process current user: " & userName

Set obj = Nothing
%>
```
