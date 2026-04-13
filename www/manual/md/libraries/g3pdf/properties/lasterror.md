# LastError

## Overview

Returns the description of the last error encountered during operations. Returns an empty string if no error occurred.

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

val = pdf.LastError
Response.Write CStr(val)

Set pdf = Nothing
%>
```
