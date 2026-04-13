# SetX

## Overview

Defines the abscissa of the current position.

## Syntax

```asp
obj.SetX x
```

## Parameters

- `x` (Double): The value of the abscissa.

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
