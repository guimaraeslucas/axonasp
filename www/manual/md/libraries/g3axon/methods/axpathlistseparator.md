# axpathlistseparator

## Overview
Retrieves the operating system's path list separator character in G3Pix AxonASP.

## Syntax
```asp
result = obj.axpathlistseparator()
```

## Parameters and Arguments
None.

## Return Values
Returns a String containing the character used by the operating system to separate multiple paths in a list (e.g., in the PATH environment variable).

## Remarks
Typically returns ";" on Windows systems and ":" on Unix-like systems.

## Code Example
```asp
<%
Option Explicit
Dim obj, listSeparator
Set obj = Server.CreateObject("G3AXON.FUNCTIONS")

listSeparator = obj.axpathlistseparator()
Response.Write "System Path List Separator: " & listSeparator

Set obj = Nothing
%>
```
