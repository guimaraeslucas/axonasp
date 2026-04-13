# axisset

## Overview

The `axisset` method checks if a value has been initialized and is neither `Empty` nor `Null`.

## Syntax

```asp
result = obj.axisset(value)
```

## Parameters and Arguments

- **value** (Variant): The value to check.

## Return Values

Returns a Boolean indicating whether the value is set. Returns `True` if the value is not `Empty` and not `Null`, otherwise `False`.

## Remarks

- This method is part of the G3Pix AxonASP library.
- It is the inverse of checking for `Empty` or `Null` using standard VBScript functions.
- Method names in G3Pix AxonASP are case-insensitive.

## Code Example

```asp
<%
Option Explicit
Dim ax, val
Set ax = Server.CreateObject("G3AXON.FUNCTIONS")

If Not ax.axisset(val) Then
    Response.Write "Variable 'val' is not set.<br>"
End If

val = "Initialized"

If ax.axisset(val) Then
    Response.Write "Variable 'val' is now set.<br>"
End If

Set ax = Nothing
%>
```
