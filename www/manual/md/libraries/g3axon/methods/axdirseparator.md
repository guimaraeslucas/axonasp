# axdirseparator

## Overview
Retrieves the operating system's directory separator character in G3Pix AxonASP.

## Syntax
```asp
result = obj.axdirseparator()
```

## Parameters and Arguments
None.

## Return Values
Returns a String containing the character used by the operating system to separate directory levels in a path.

## Remarks
Typically returns "\" on Windows systems and "/" on Unix-like systems.

## Code Example
```asp
<%
Option Explicit
Dim obj, separator
Set obj = Server.CreateObject("G3AXON.FUNCTIONS")

separator = obj.axdirseparator()
Response.Write "System Directory Separator: " & separator

Set obj = Nothing
%>
```
