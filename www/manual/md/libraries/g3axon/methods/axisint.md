# axisint

## Overview

The `axisint` method checks if the internal Virtual Machine (VM) type of a value is an Integer.

## Syntax

```asp
result = obj.axisint(value)
```

## Parameters and Arguments

- **value** (Variant): The value to check.

## Return Values

Returns a Boolean indicating whether the internal VM type of the value is `VTInteger`. Returns `True` if it is an Integer, otherwise `False`.

## Remarks

- This method is part of the G3Pix AxonASP library.
- It checks the native VM type, which may differ from standard VBScript `VarType` in some edge cases.
- Method names in G3Pix AxonASP are case-insensitive.

## Code Example

```asp
<%
Option Explicit
Dim ax, val
Set ax = Server.CreateObject("G3AXON.FUNCTIONS")

val = 100

If ax.axisint(val) Then
    Response.Write "Value is an Integer."
Else
    Response.Write "Value is not an Integer."
End If

Set ax = Nothing
%>
```
