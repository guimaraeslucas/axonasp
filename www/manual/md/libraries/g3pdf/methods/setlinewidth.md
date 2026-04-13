# SetLineWidth

## Overview

Defines the line width.

## Syntax

```asp
obj.SetLineWidth width
```

## Parameters

- `width` (Double): The width.

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
