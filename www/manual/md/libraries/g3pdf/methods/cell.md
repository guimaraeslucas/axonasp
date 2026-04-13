# Cell

## Overview

Prints a cell (rectangular area) with optional borders, background color, and character string. The upper-left corner of the cell corresponds to the current position.

## Syntax

```asp
obj.Cell w, h, txt, [border], [ln], [align], [fill], [link]
```

## Parameters

- `w` (Double): Cell width. If 0, the cell extends up to the right margin.
- `h` (Double): Cell height.
- `txt` (String): String to print.
- `border` (Variant, Optional): Indicates if borders must be drawn (0, 1, or string like `LRTB`).
- `ln` (Integer, Optional): Indicates where the current position should go after the call (0: to the right, 1: to the beginning of the next line, 2: below).
- `align` (String, Optional): Text alignment (`L`, `C`, `R`).
- `fill` (Boolean, Optional): Indicates if the cell background must be painted.
- `link` (Integer, Optional): Identifier or string for a link.

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
