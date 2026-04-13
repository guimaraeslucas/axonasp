# SetTopMargin

## Overview

Defines the top margin.

## Syntax

```asp
obj.SetTopMargin margin
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
