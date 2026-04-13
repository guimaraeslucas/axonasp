# axmax

## Overview
Returns the largest numeric value from the provided arguments in G3Pix AxonASP.

## Syntax
```asp
result = obj.axmax(n1, n2, ..., nN)
```

## Parameters and Arguments
- **n1, n2, ..., nN** (Numeric): A variable number of numeric values to be compared.

## Return Values
Returns a Double representing the maximum value found among all arguments. If no arguments are provided, it returns 0.

## Remarks
The function automatically converts non-numeric values to their numeric equivalent before comparison.

## Code Example
```asp
<%
Option Explicit
Dim obj, maxVal
Set obj = Server.CreateObject("G3AXON.FUNCTIONS")

' Returns 45.7
maxVal = obj.axmax(10, 45.7, 32, -5)
Response.Write "Max value: " & maxVal

Set obj = Nothing
%>
```
