# axhostnamevalue

## Overview
Retrieves the hostname of the current machine in G3Pix AxonASP.

## Syntax
```asp
result = obj.axhostnamevalue()
```

## Parameters and Arguments
None.

## Return Values
Returns a String containing the machine's hostname.

## Remarks
If the hostname cannot be determined, an empty string is returned.

## Code Example
```asp
<%
Option Explicit
Dim obj, hostName
Set obj = Server.CreateObject("G3AXON.FUNCTIONS")

hostName = obj.axhostnamevalue()
Response.Write "Server Hostname: " & hostName

Set obj = Nothing
%>
```
