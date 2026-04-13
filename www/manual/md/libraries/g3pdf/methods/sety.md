# SetY

## Overview

Moves the current abscissa back to the left margin and sets the ordinate.

## Syntax

```asp
obj.SetY y, [resetX]
```

## Parameters

- `y` (Double): The value of the ordinate.
- `resetX` (Boolean, Optional): Indicates if the abscissa should be reset to the left margin. Default is true.

## Return Value

**Returns:** Empty

## Code Example

```asp
<%
Option Explicit

Dim pdf
Set pdf = Server.CreateObject("G3PDF")

pdf.Reset "P", "mm", "A4"
pdf.AddPage

' Perform method operations here

Set pdf = Nothing
%>
```
