# axplatformbits

## Overview
Retrieves the platform architecture bits for the G3Pix AxonASP runtime.

## Syntax
```asp
result = obj.axplatformbits()
```

## Parameters and Arguments
None.

## Return Values
Returns an Integer representing the platform architecture (32 or 64).

## Remarks
This indicates the word size of the current operating system and process architecture.

## Code Example
```asp
<%
Option Explicit
Dim obj, bits
Set obj = Server.CreateObject("G3AXON.FUNCTIONS")

bits = obj.axplatformbits()
Response.Write "Platform architecture: " & bits & "-bit"

Set obj = Nothing
%>
```
