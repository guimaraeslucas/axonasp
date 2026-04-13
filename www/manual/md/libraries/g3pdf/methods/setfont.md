# SetFont

## Overview

Sets the font used to print character strings.

## Syntax

```asp
obj.SetFont family, [style], [size]
```

## Parameters

- `family` (String): Family font. It can be a standard TrueType/Type1 font.
- `style` (String, Optional): Font style. Possible values are empty string (regular), `B` (bold), `I` (italic), `U` (underline), or any combination.
- `size` (Double, Optional): Font size in points.

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
