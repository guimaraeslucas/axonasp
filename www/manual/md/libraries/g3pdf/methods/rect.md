# Rect

## Overview

Outputs a rectangle.

## Syntax

```asp
obj.Rect x, y, w, h, [style]
```

## Parameters

- `x` (Double): Abscissa of upper-left corner.
- `y` (Double): Ordinate of upper-left corner.
- `w` (Double): Width.
- `h` (Double): Height.
- `style` (String, Optional): Style of rendering. Possible values: `D` (Draw), `F` (Fill), `DF` or `FD` (Draw and fill).

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
