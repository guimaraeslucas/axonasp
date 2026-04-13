# SetXY

## Overview

Defines the abscissa and ordinate of the current position.

## Syntax

```asp
obj.SetXY x, y
```

## Parameters

- `x` (Double): The value of the abscissa.
- `y` (Double): The value of the ordinate.

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
