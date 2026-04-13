# SetDrawColor

## Overview

Defines the color used for all drawing operations (lines, rectangles and cell borders). It can be expressed in RGB components or grayscale.

## Syntax

```asp
obj.SetDrawColor r, [g], [b]
```

## Parameters

- `r` (Integer): If `g` and `b` are given, this indicates the red component; else, it indicates the grayscale level.
- `g` (Integer, Optional): Green component (between 0 and 255).
- `b` (Integer, Optional): Blue component (between 0 and 255).

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
