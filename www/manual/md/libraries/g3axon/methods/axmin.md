# axmin

## Overview
Returns the smallest numeric value from the provided arguments in G3Pix AxonASP.

## Syntax
```asp
result = obj.axmin(n1, n2, ..., nN)
```

## Parameters and Arguments
- **n1, n2, ..., nN** (Numeric): A variable number of numeric values to be compared.

## Return Values
Returns a Double representing the minimum value found among all arguments. If no arguments are provided, it returns 0.

## Remarks
The function automatically converts non-numeric values to their numeric equivalent before comparison.

## Code Example
```asp
<%
Option Explicit
Dim obj, minVal
Set obj = Server.CreateObject("G3AXON.FUNCTIONS")

' Returns -5.2
minVal = obj.axmin(10, 45, 32, -5.2)
Response.Write "Min value: " & minVal

Set obj = Nothing
%>
```
