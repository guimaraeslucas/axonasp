## PDF_LIB (G3PDF) Library Implementation Summary

### Overview
`PDF_LIB` is implemented as the `G3PDF` component and exposed through `Server.CreateObject("G3PDF")` with compatibility aliases `PDF` and `FPDF`.
The implementation is a Go translation of FPDF 1.86 adapted to AxonASP execution context, with support for direct PDF generation, image insertion, metadata, links, and HTML-to-PDF rendering.

### Files

#### Core Library
1. **`server/pdf_lib.go`**
   - Main PDF engine and VBScript method dispatcher.
   - Page management, text rendering, drawing primitives, images, links, metadata, output modes.
   - HTML rendering support (`WriteHTML`, `WriteHTMLFile`) translated from `extra.php` behavior.

2. **`server/pdf_fpdf_font_data.go`**
   - Go-native translated core FPDF font metrics and Unicode maps.
   - Removes runtime dependency on PHP font definition parsing.

#### Runtime Integration
1. **`server/executor.go`**
   - `CreateObject` mappings:
     - `G3PDF`
     - `PDF`
     - `FPDF`

### Object Creation

```vbscript
Dim pdf
Set pdf = Server.CreateObject("G3PDF")
```

### Main Methods

#### Document and Page
- `New([orientation], [unit], [size])` / `Reset(...)`
- `AddPage([orientation], [size], [rotation])`
- `Close()`
- `Output([dest], [fileName], [isUTF8])`

#### Text and Layout
- `SetFont(family, [style], [size])`
- `SetFontSize(size)`
- `Cell(w, h, text, [border], [ln], [align], [fill], [link])`
- `MultiCell(w, h, text, [border], [align], [fill])`
- `Write(h, text, [link])`
- `Text(x, y, text)`
- `Ln([h])`

#### Graphics
- `Line(x1, y1, x2, y2)`
- `Rect(x, y, w, h, [style])`
- `SetLineWidth(width)`
- `SetDrawColor(r, [g], [b])`
- `SetFillColor(r, [g], [b])`
- `SetTextColor(r, [g], [b])`

#### Images
- `Image(path, [x], [y], [w], [h], [type], [link])`

#### Links
- `AddLink()`
- `SetLink(linkId, [y], [page])`
- `Link(x, y, w, h, linkOrUrl)`

#### Metadata and Output Behavior
- `SetTitle(text, [isUTF8])`
- `SetAuthor(text, [isUTF8])`
- `SetSubject(text, [isUTF8])`
- `SetKeywords(text, [isUTF8])`
- `SetCreator(text, [isUTF8])`
- `AliasNbPages([alias])`
- `SetDisplayMode([zoom], [layout])`
- `SetCompression(enabled)`

#### HTML Rendering
- `WriteHTML(htmlString)`
- `WriteHTMLFile(path)`
- Alias methods in dispatcher:
  - `html` => `WriteHTML`
  - `writehtml` => `WriteHTML`
  - `htmlfile` / `writehtmlfile` / `loadhtmlfile` => `WriteHTMLFile`

### Output Destinations
- `Output("I")`: inline output to HTTP response.
- `Output("D", "file.pdf")`: attachment download response.
- `Output("F", "path.pdf")`: save file in server filesystem.
- `Output("S")`: returns PDF bytes/string.

### HTML Support Notes
The HTML renderer supports common tags used by classic FPDF extra parser behavior, including:
- Typography: `B`, `I`, `U`, `STRONG`, `EM`, `FONT`
- Blocks: `P`, `DIV`, `CENTER`, `BR`, headings `H1..H6`, `PRE`, `BLOCKQUOTE`, `HR`
- Lists: `UL`, `OL`, `LI`, `DL`, `DT`, `DD`
- Links and media: `A`, `IMG`
- Basic table flow: `TABLE`, `TR`, `TD`

### Usage Examples

#### 1) Basic PDF
```vbscript
<%
Dim pdf
Set pdf = Server.CreateObject("G3PDF")

pdf.AddPage
pdf.SetFont "helvetica", "B", 16
pdf.Cell 0, 12, "Hello from AxonASP PDF_LIB", 0, 1, "L", False, ""

Response.ContentType = "application/pdf"
pdf.Output "I", "basic.pdf", True
%>
```

#### 2) PDF with Image
```vbscript
<%
Dim pdf
Set pdf = Server.CreateObject("G3PDF")

pdf.AddPage
pdf.SetFont "helvetica", "", 12
pdf.Cell 0, 8, "Image test", 0, 1, "L", False, ""
pdf.Image Server.MapPath("/asplite-test/uploads/Png.png"), 20, 30, 70, 0, "", ""

Response.ContentType = "application/pdf"
pdf.Output "I", "image.pdf", True
%>
```

#### 3) PDF from HTML String
```vbscript
<%
Dim pdf, html
Set pdf = Server.CreateObject("G3PDF")

html = "<h2>HTML to PDF</h2><p><b>Bold</b> and <i>italic</i> sample.</p>"

pdf.AddPage
pdf.SetFont "helvetica", "", 12
pdf.WriteHTML html

Response.ContentType = "application/pdf"
pdf.Output "I", "html.pdf", True
%>
```

### Test Pages
Validation pages are available in `www/tests/`:
- `test_pdf_basic.asp`
- `test_pdf_image.asp`
- `test_pdf_html.asp`
- `test_pdf_html_sample.html`
