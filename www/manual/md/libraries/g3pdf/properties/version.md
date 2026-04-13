# Version

## Overview

Returns the current version string of the underlying go-pdf/fpdf library wrapper.

## Property Type

**Returns:** String

## Code Example

```asp
<%
Option Explicit

Dim pdf, val
Set pdf = Server.CreateObject("G3PDF")

pdf.Reset "P", "mm", "A4"
pdf.AddPage

val = pdf.Version
Response.Write CStr(val)

Set pdf = Nothing
%>
```
