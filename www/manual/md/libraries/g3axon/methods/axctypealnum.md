# axctypealnum

## Overview

The `axctypealnum` method checks if a given string consists entirely of alphanumeric characters (a-z, A-Z, and 0-9).

## Syntax

```asp
result = obj.axctypealnum(inputString)
```

## Parameters and Arguments

- **inputString** (String): The string to check for alphanumeric characters.

## Return Values

Returns a Boolean indicating whether the string contains only alphanumeric characters. Returns `True` if all characters are alphanumeric, otherwise `False`. It also returns `False` for empty strings.

## Remarks

- This method is part of the G3Pix AxonASP library.
- It only considers ASCII alphanumeric characters.
- Method names in G3Pix AxonASP are case-insensitive.

## Code Example

```asp
<%
Option Explicit
Dim ax, str1, str2
Set ax = Server.CreateObject("G3AXON.FUNCTIONS")

str1 = "AxonASP2026"
str2 = "Axon-ASP"

If ax.axctypealnum(str1) Then
    Response.Write str1 & " is alphanumeric.<br>"
End If

If Not ax.axctypealnum(str2) Then
    Response.Write str2 & " contains non-alphanumeric characters.<br>"
End If

Set ax = Nothing
%>
```
