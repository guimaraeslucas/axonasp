# axsysteminfo

## Overview
Retrieves various system and runtime environment information in G3Pix AxonASP.

## Syntax
```asp
result = obj.axsysteminfo([mode])
```

## Parameters and Arguments
- **mode** (String, Optional): A single character specifying the type of information to return:
  - "s": Operating system name (e.g., "windows", "linux").
  - "n": Hostname of the machine.
  - "v": Go runtime version.
  - "m": Machine architecture (e.g., "amd64", "arm64").
  - "a": All available information (default).

## Return Values
Returns a String containing the requested system information. If no mode is provided, it defaults to returning all information in a single string.

## Remarks
If an invalid or empty mode is provided, the function defaults to "a".

## Code Example
```asp
<%
Option Explicit
Dim obj
Set obj = Server.CreateObject("G3AXON.FUNCTIONS")

Response.Write "OS: " & obj.axsysteminfo("s") & "<br>"
Response.Write "Host: " & obj.axsysteminfo("n") & "<br>"
Response.Write "Architecture: " & obj.axsysteminfo("m") & "<br>"
Response.Write "Full Info: " & obj.axsysteminfo("a")

Set obj = Nothing
%>
```
