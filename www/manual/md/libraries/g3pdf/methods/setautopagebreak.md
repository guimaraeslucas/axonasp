# SetAutoPageBreak

## Overview

Enables or disables the automatic page breaking mode.

## Syntax

```asp
obj.SetAutoPageBreak auto, [margin]
```

## Parameters

- `auto` (Boolean): Boolean indicating if mode should be on or off.
- `margin` (Double, Optional): Distance from the bottom of the page that triggers the break.

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
