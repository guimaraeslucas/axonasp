# SetTextColor

## Overview

Defines the color used for text. It can be expressed in RGB components or grayscale.

## Syntax

```asp
obj.SetTextColor r, [g], [b]
```

## Parameters

- `r` (Integer): Red component or grayscale level.
- `g` (Integer, Optional): Green component.
- `b` (Integer, Optional): Blue component.

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
