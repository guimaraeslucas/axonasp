# Write

## Overview

Prints text from the current position. When the right margin is reached (or the newline character is met), a line break occurs and text continues from the left margin.

## Syntax

```asp
obj.Write h, txt, [link]
```

## Parameters

- `h` (Double): Line height.
- `txt` (String): String to print.
- `link` (Variant, Optional): URL or identifier.

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
