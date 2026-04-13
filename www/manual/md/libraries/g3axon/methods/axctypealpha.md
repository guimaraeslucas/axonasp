# axctypealpha

## Overview

The `axctypealpha` method checks if a given string consists entirely of alphabetic characters (a-z and A-Z).

## Syntax

```asp
result = obj.axctypealpha(inputString)
```

## Parameters and Arguments

- **inputString** (String): The string to check for alphabetic characters.

## Return Values

Returns a Boolean indicating whether the string contains only alphabetic characters. Returns `True` if all characters are alphabetic, otherwise `False`. It also returns `False` for empty strings.

## Remarks

- This method is part of the G3Pix AxonASP library.
- It only considers ASCII alphabetic characters.
- Method names in G3Pix AxonASP are case-insensitive.

## Code Example

```asp
<%
Option Explicit
Dim ax, str1, str2
Set ax = Server.CreateObject("G3AXON.FUNCTIONS")

str1 = "AxonASP"
str2 = "AxonASP 2.0"

If ax.axctypealpha(str1) Then
    Response.Write str1 & " is alphabetic.<br>"
End If

If Not ax.axctypealpha(str2) Then
    Response.Write str2 & " contains non-alphabetic characters.<br>"
End If

Set ax = Nothing
%>
```
