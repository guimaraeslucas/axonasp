# AddPage

## Overview

Adds a new page to the PDF document. Must be called before outputting any content to a page.

## Syntax

```asp
obj.AddPage [orientation], [size]
```

## Parameters

- `orientation` (String, Optional): Document orientation. Possible values are `P` (Portrait) or `L` (Landscape).
- `size` (String, Optional): Page format (e.g., `A4`, `Letter`).

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
