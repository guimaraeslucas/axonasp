# MultiCell

## Overview

Allows printing text with line breaks. They can be automatic (as soon as the text reaches the right border of the cell) or explicit (via the newline character).

## Syntax

```asp
obj.MultiCell w, h, txt, [border], [align], [fill]
```

## Parameters

- `w` (Double): Width of cells. If 0, they extend up to the right margin of the page.
- `h` (Double): Height of cells.
- `txt` (String): String to print.
- `border` (Variant, Optional): Indicates if borders must be drawn (0, 1, or string like `LRTB`).
- `align` (String, Optional): Allows to center or align the text. Possible values: `L`, `C`, `R`, `J`.
- `fill` (Boolean, Optional): Indicates if the cell background must be painted.

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
