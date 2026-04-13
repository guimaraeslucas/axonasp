# SetLeftMargin

## Overview

Defines the left margin.

## Syntax

```asp
obj.SetLeftMargin margin
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
