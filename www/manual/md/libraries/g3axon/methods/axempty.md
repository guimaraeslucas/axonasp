# axempty

## Overview

The `axempty` method checks if a value is considered "empty" according to extended criteria. This includes traditional `Empty` and `Null` values, as well as zero-like values and empty strings.

## Syntax

```asp
result = obj.axempty(value)
```

## Parameters and Arguments

- **value** (Variant): The value to check for emptiness.

## Return Values

Returns a Boolean indicating whether the value is considered empty. Returns `True` if the value is:
- `Empty` or `Null`.
- An empty string (`""`).
- The integer `0`.
- The double `0.0`.
- The boolean `False`.

Otherwise, it returns `False`.

## Remarks

- This method is part of the G3Pix AxonASP library.
- It provides a convenient way to check for various "no-value" states in a single call.
- Method names in G3Pix AxonASP are case-insensitive.

## Code Example

```asp
<%
Option Explicit
Dim ax, val1, val2
Set ax = Server.CreateObject("G3AXON.FUNCTIONS")

val1 = ""
val2 = 123

If ax.axempty(val1) Then
    Response.Write "val1 is empty.<br>"
End If

If Not ax.axempty(val2) Then
    Response.Write "val2 is not empty.<br>"
End If

Set ax = Nothing
%>
```
