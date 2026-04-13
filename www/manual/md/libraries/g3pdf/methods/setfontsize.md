# SetFontSize

## Overview

Defines the size of the current font.

## Syntax

```asp
obj.SetFontSize size
```

## Parameters

- `size` (Double): The size (in points).

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
