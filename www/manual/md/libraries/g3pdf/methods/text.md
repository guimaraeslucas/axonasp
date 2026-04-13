# Text

## Overview

Prints a character string. The origin is on the left of the first character, on the baseline.

## Syntax

```asp
obj.Text x, y, txt
```

## Parameters

- `x` (Double): Abscissa of the origin.
- `y` (Double): Ordinate of the origin.
- `txt` (String): String to print.

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
