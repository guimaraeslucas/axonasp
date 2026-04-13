# Image

## Overview

Outputs an image into the document.

## Syntax

```asp
obj.Image imageName, x, y, [w], [h], [type], [link]
```

## Parameters

- `imageName` (String): Path or URL of the image.
- `x` (Double): Abscissa of the upper-left corner.
- `y` (Double): Ordinate of the upper-left corner.
- `w` (Double, Optional): Width of the image in the page.
- `h` (Double, Optional): Height of the image in the page.
- `type` (String, Optional): Image format (e.g., `JPG`, `PNG`, `GIF`).
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
