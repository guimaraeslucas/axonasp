# axprocessid

## Overview
Retrieves the process identifier (PID) of the current G3Pix AxonASP process.

## Syntax
```asp
result = obj.axprocessid()
```

## Parameters and Arguments
None.

## Return Values
Returns an Integer representing the current process identifier.

## Remarks
The process ID is unique to the current running instance of the application.

## Code Example
```asp
<%
Option Explicit
Dim obj, pid
Set obj = Server.CreateObject("G3AXON.FUNCTIONS")

pid = obj.axprocessid()
Response.Write "Current Process ID (PID): " & pid

Set obj = Nothing
%>
```
