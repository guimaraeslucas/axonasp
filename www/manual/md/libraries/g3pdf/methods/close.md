# Close

## Overview

Terminates the PDF document. It is not necessary to call this method explicitly for most outputs as it is called automatically by output methods.

## Syntax

```asp
obj.Close
```

## Parameters

None.

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
