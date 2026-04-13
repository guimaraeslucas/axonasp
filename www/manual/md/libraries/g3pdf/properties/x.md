# X

## Overview

Gets the abscissa of the current position.

## Property Type

**Returns:** Double

## Code Example

```asp
<%
Option Explicit

Dim pdf, val
Set pdf = Server.CreateObject("G3PDF")

pdf.Reset "P", "mm", "A4"
pdf.AddPage

val = pdf.X
Response.Write CStr(val)

Set pdf = Nothing
%>
```
