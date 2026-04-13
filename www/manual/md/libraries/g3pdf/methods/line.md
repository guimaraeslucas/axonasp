# Line

## Overview

Draws a line between two points.

## Syntax

```asp
obj.Line x1, y1, x2, y2
```

## Parameters

- `x1` (Double): Abscissa of first point.
- `y1` (Double): Ordinate of first point.
- `x2` (Double): Abscissa of second point.
- `y2` (Double): Ordinate of second point.

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
