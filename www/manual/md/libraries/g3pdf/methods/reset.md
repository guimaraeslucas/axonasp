# Reset

## Overview

Resets the PDF engine, clearing all content, settings, and pages. It optionally allows setting the new defaults for the document.

## Syntax

```asp
obj.Reset [orientation], [unit], [size]
```

## Parameters

- `orientation` (String, Optional): Default page orientation (`P` or `L`).
- `unit` (String, Optional): User unit of measure (e.g., `pt`, `mm`, `cm`, `in`).
- `size` (String, Optional): Default page format (e.g., `A4`, `Letter`).

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
