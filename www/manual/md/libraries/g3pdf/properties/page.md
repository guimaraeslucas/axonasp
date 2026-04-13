# Page

## Overview

Returns the current page number.

## Property Type

**Returns:** Integer

## Code Example

```asp
<%
Option Explicit

Dim pdf, val
Set pdf = Server.CreateObject("G3PDF")

pdf.Reset "P", "mm", "A4"
pdf.AddPage

val = pdf.Page
Response.Write CStr(val)

Set pdf = Nothing
%>
```
