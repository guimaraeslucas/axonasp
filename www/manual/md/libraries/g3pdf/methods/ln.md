# Ln

## Overview

Performs a line break. The current abscissa goes back to the left margin and the ordinate increases by the amount passed in parameter.

## Syntax

```asp
obj.Ln [h]
```

## Parameters

- `h` (Double, Optional): The height of the break. By default, the value equals the height of the last printed cell.

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
