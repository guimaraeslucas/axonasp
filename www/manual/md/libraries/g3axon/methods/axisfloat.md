# axisfloat

## Overview

The `axisfloat` method checks if the internal Virtual Machine (VM) type of a value is a Double (floating-point number).

## Syntax

```asp
result = obj.axisfloat(value)
```

## Parameters and Arguments

- **value** (Variant): The value to check.

## Return Values

Returns a Boolean indicating whether the internal VM type of the value is `VTDouble`. Returns `True` if it is a floating-point number, otherwise `False`.

## Remarks

- This method is part of the G3Pix AxonASP library.
- It checks the native VM type, which corresponds to the Double precision floating-point type.
- Method names in G3Pix AxonASP are case-insensitive.

## Code Example

```asp
<%
Option Explicit
Dim ax, val
Set ax = Server.CreateObject("G3AXON.FUNCTIONS")

val = 123.45

If ax.axisfloat(val) Then
    Response.Write "Value is a Float (Double)."
Else
    Response.Write "Value is not a Float (Double)."
End If

Set ax = Nothing
%>
```
