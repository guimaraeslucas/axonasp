# Create PDF documents using G3PDF

## Overview
The `G3PDF` object is a high-performance native library for generating PDF documents directly from Classic ASP code. It provides functions to add pages, format text, draw shapes, and manage document layout.

## Syntax
```asp
Dim pdf
Set pdf = Server.CreateObject("G3PDF")
```

## Parameters
None for instantiation.

## Return Values
Returns a native `G3PDF` object that can be used to construct PDF documents.

## Remarks
- Requires `Server.CreateObject("G3PDF")`. No aliases are supported.
- Powered by `go-pdf/fpdf` optimized for AxonASP.
- Ensure that memory and binary output streams are managed correctly.
- Call `Close` when finished applying content to the PDF object.

## Code Example
```asp
<%
Dim pdf
Set pdf = Server.CreateObject("G3PDF")

pdf.AddPage "", "", 0
pdf.SetFont "Arial", "B", 16
pdf.Cell 40, 10, "Hello World!", 1, 0, "C", False, ""

' Further output code...
%>
```

