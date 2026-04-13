# SetMargins

## Overview

Defines the left, top, and right margins. By default, they equal 1 cm.

## Syntax

```asp
obj.SetMargins left, top, [right]
```

## Parameters

- `left` (Double): Left margin.
- `top` (Double): Top margin.
- `right` (Double, Optional): Right margin. Default to the value of the left margin.

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
