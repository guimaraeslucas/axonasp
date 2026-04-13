# SetRightMargin

## Overview

Defines the right margin.

## Syntax

```asp
obj.SetRightMargin margin
```

## Parameters

- `margin` (Double): The margin.

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
