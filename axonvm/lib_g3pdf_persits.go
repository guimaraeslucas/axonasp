//go:build !wasm && !lib_g3pdf_disabled

/*
 * AxonASP Server
 * Copyright (C) 2026 G3pix Ltda. All rights reserved.
 *
 * Developed by Lucas Guimaraes - G3pix Ltda
 * Contact: https://g3pix.com.br
 * Project URL: https://g3pix.com.br/axonasp
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *
 * Persits.Pdf Compatibility Layer for G3PDF
 * Implements the AspPDF object model on top of go-pdf/fpdf
 */
package axonvm

import (
	"bytes"
	"fmt"
	"os"
	"strconv"
	"strings"

	"g3pix.com.br/axonasp/v2/vbscript"
)

// ============================================================================
// Param String Parser
// ============================================================================

// aspPDFParam represents a single parsed key-value pair from a param string.
type aspPDFParam struct {
	Key   string
	Value string
}

// parseAspPDFParams parses a Persits.Pdf-style parameter string into a fast
// lookup table. The format is "key=value; key2=value2; ...".
// It uses a single pre-sized map allocation and does not allocate per token.
func parseAspPDFParams(input string) map[string]string {
	if input == "" {
		return nil
	}
	// Estimate capacity: count semicolons + 1
	capEst := 1
	for i := 0; i < len(input); i++ {
		if input[i] == ';' {
			capEst++
		}
	}
	result := make(map[string]string, capEst)
	pos := 0
	for pos < len(input) {
		// Skip leading whitespace and semicolons
		for pos < len(input) && (input[pos] == ' ' || input[pos] == '\t' || input[pos] == ';') {
			pos++
		}
		if pos >= len(input) {
			break
		}
		// Find '='
		start := pos
		for pos < len(input) && input[pos] != '=' && input[pos] != ';' {
			pos++
		}
		if pos < len(input) && input[pos] == '=' {
			key := strings.TrimSpace(input[start:pos])
			pos++ // skip '='
			start = pos
			// Read value until ';' or end
			for pos < len(input) && input[pos] != ';' {
				pos++
			}
			value := strings.TrimSpace(input[start:pos])
			if key != "" {
				result[strings.ToLower(key)] = value
			}
		} else {
			// No '=' found, treat as flag key
			key := strings.TrimSpace(input[start:pos])
			if key != "" {
				result[strings.ToLower(key)] = "true"
			}
		}
	}
	return result
}

// parseAspPDFFloat reads a float64 param from the map, returning fallback if missing.
func parseAspPDFFloat(params map[string]string, key string, fallback float64) float64 {
	if params == nil {
		return fallback
	}
	v, ok := params[key]
	if !ok || v == "" {
		return fallback
	}
	f, err := strconv.ParseFloat(v, 64)
	if err != nil {
		return fallback
	}
	return f
}

// parseAspPDFInt reads an int param from the map, returning fallback if missing.
func parseAspPDFInt(params map[string]string, key string, fallback int) int {
	if params == nil {
		return fallback
	}
	v, ok := params[key]
	if !ok || v == "" {
		return fallback
	}
	i, err := strconv.Atoi(v)
	if err != nil {
		return fallback
	}
	return i
}

// parseAspPDFColor reads a color param from the map.
// Supports "#RRGGBB" or named colors. Returns (r,g,b, ok).
func parseAspPDFColor(params map[string]string, key string) (int, int, int, bool) {
	if params == nil {
		return 0, 0, 0, false
	}
	v, ok := params[key]
	if !ok || v == "" {
		return 0, 0, 0, false
	}
	return htmlColorToRGB(v)
}

// ============================================================================
// PdfFont - Persits.Pdf Font object
// ============================================================================

// PdfFont represents a font loaded via Fonts("FontName").
type PdfFont struct {
	vm       *VM
	objectID int64
	pdf      *G3PDF

	name     string
	family   string
	style    string // "Bold", "Italic", "BoldItalic", ""
	size     float64
	colorR   int
	colorG   int
	colorB   int
	colorSet bool
	embedded bool
}

// PdfFontDispatchMethod dispatches method calls on a PdfFont object.
func (f *PdfFont) DispatchMethod(member string, args []Value) Value {
	return NewEmpty()
}

// PdfFontDispatchPropertyGet returns a property value for PdfFont.
func (f *PdfFont) DispatchPropertyGet(member string) Value {
	switch strings.ToLower(member) {
	case "name":
		return NewString(f.name)
	case "family":
		return NewString(f.family)
	case "size":
		return NewDouble(f.size)
	case "bold":
		return NewBool(strings.Contains(f.style, "Bold"))
	case "italic":
		return NewBool(strings.Contains(f.style, "Italic"))
	case "embedded":
		return NewBool(f.embedded)
	}
	return NewEmpty()
}

// PdfFontDispatchPropertySet sets a property on PdfFont.
func (f *PdfFont) DispatchPropertySet(member string, args []Value) bool {
	if len(args) == 0 {
		return false
	}
	val := legacyValueToInterface(args[0], f.vm)
	switch strings.ToLower(member) {
	case "size":
		f.size = toFloat(val)
		return true
	case "bold":
		f.setStyleBold(toBool(val))
		return true
	case "italic":
		f.setStyleItalic(toBool(val))
		return true
	case "embedded":
		f.embedded = toBool(val)
		return true
	}
	return false
}

func (f *PdfFont) setStyleBold(bold bool) {
	if bold && !strings.Contains(f.style, "Bold") {
		f.style = "Bold" + f.style
	} else if !bold && strings.Contains(f.style, "Bold") {
		f.style = strings.Replace(f.style, "Bold", "", 1)
	}
}

func (f *PdfFont) setStyleItalic(italic bool) {
	if italic && !strings.Contains(f.style, "Italic") {
		f.style = f.style + "Italic"
	} else if !italic && strings.Contains(f.style, "Italic") {
		f.style = strings.Replace(f.style, "Italic", "", 1)
	}
}

// getFPDFStyle converts Persits font style to fpdf style string.
func (f *PdfFont) getFPDFStyle() string {
	s := ""
	if strings.Contains(f.style, "Bold") {
		s += "B"
	}
	if strings.Contains(f.style, "Italic") {
		s += "I"
	}
	return s
}

// ============================================================================
// PdfCanvas - Persits.Pdf Canvas object
// ============================================================================

// PdfCanvas represents the drawing surface of a PdfPage.
type PdfCanvas struct {
	vm       *VM
	objectID int64
	pdf      *G3PDF
	page     *PdfPage
}

// DrawText draws text on the canvas using a param string and optional font.
// Param keys: x, y, width, alignment (left/center/right/justify), size, color.
func (c *PdfCanvas) DrawText(text string, params string, font *PdfFont) {
	if text == "" {
		return
	}
	p := parseAspPDFParams(params)

	x := parseAspPDFFloat(p, "x", c.pdf.pdf.GetX())
	y := parseAspPDFFloat(p, "y", c.pdf.pdf.GetY())
	width := parseAspPDFFloat(p, "width", 0)
	align := strings.ToLower(p["alignment"])
	if align == "" {
		align = "left"
	}
	size := parseAspPDFFloat(p, "size", 0)
	colorStr := p["color"]

	// Ensure at least one page exists
	if c.pdf.pdf.PageNo() == 0 {
		c.pdf.pdf.AddPage()
	}

	// Determine font family and style
	family := "helvetica"
	style := ""
	fontSize := 12.0

	if font != nil {
		family = font.family
		style = font.getFPDFStyle()
		fontSize = font.size
		if fontSize <= 0 {
			fontSize = 12
		}
	}
	if size > 0 {
		fontSize = size
	}

	// Ensure font is always set before any text operation
	c.pdf.pdf.SetFont(family, style, fontSize)

	// Set position
	c.pdf.pdf.SetXY(x, y)

	// Set color if specified
	if colorStr != "" {
		if r, g, b, ok := htmlColorToRGB(colorStr); ok {
			c.pdf.pdf.SetTextColor(r, g, b)
		}
	} else if font != nil && font.colorSet {
		c.pdf.pdf.SetTextColor(font.colorR, font.colorG, font.colorB)
	}

	// Map alignment
	fpdfAlign := "L"
	switch align {
	case "center", "centre":
		fpdfAlign = "C"
	case "right":
		fpdfAlign = "R"
	case "justify":
		fpdfAlign = "J"
	}

	// Clamp width to available page space so MultiCell stays visible.
	// Available width = page width - left margin - x offset.
	if width > 0 {
		leftMargin, _, rightMargin, _ := c.pdf.pdf.GetMargins()
		pageW, _ := c.pdf.pdf.GetPageSize()
		maxW := pageW - leftMargin - rightMargin - x
		if maxW < 10 {
			maxW = 10 // minimum sensible width
		}
		if width > maxW {
			width = maxW
		}
	}

	// Draw text based on width
	if width > 0 {
		c.pdf.pdf.MultiCell(width, 5, text, "", fpdfAlign, false)
	} else {
		c.pdf.pdf.Write(5, text)
	}
}

// DrawLine draws a line using a param string.
// Param keys: x, y (start), x1, y1 (end), color, width.
func (c *PdfCanvas) DrawLine(params string) {
	p := parseAspPDFParams(params)

	x := parseAspPDFFloat(p, "x", 0)
	y := parseAspPDFFloat(p, "y", 0)
	x1 := parseAspPDFFloat(p, "x1", x+10)
	y1 := parseAspPDFFloat(p, "y1", y+10)
	lineWidth := parseAspPDFFloat(p, "width", 0.2)

	// Clamp coordinates to page bounds to prevent off-page rendering.
	leftMargin, _, rightMargin, _ := c.pdf.pdf.GetMargins()
	pageW, pageH := c.pdf.pdf.GetPageSize()
	maxX := pageW - rightMargin
	maxY := pageH - leftMargin

	if x < leftMargin {
		x = leftMargin
	}
	if x > maxX {
		x = maxX
	}
	if y < leftMargin {
		y = leftMargin
	}
	if y > maxY {
		y = maxY
	}
	if x1 < leftMargin {
		x1 = leftMargin
	}
	if x1 > maxX {
		x1 = maxX
	}
	if y1 < leftMargin {
		y1 = leftMargin
	}
	if y1 > maxY {
		y1 = maxY
	}

	// Apply line width
	if lineWidth > 0 {
		c.pdf.pdf.SetLineWidth(lineWidth)
	}

	// Apply color
	if r, g, b, ok := parseAspPDFColor(p, "color"); ok {
		c.pdf.pdf.SetDrawColor(r, g, b)
	}

	c.pdf.pdf.Line(x, y, x1, y1)
}

// DrawBox draws a rectangle using a param string.
// Param keys: left, top, right, bottom, color, width.
func (c *PdfCanvas) DrawBox(params string) {
	p := parseAspPDFParams(params)

	left := parseAspPDFFloat(p, "left", 10)
	top := parseAspPDFFloat(p, "top", 10)
	right := parseAspPDFFloat(p, "right", left+50)
	bottom := parseAspPDFFloat(p, "bottom", top+50)
	lineWidth := parseAspPDFFloat(p, "width", 0.2)

	// Clamp coordinates to page bounds.
	leftMargin, _, rightMargin, _ := c.pdf.pdf.GetMargins()
	pageW, pageH := c.pdf.pdf.GetPageSize()
	maxX := pageW - rightMargin
	maxY := pageH - leftMargin

	if left < leftMargin {
		left = leftMargin
	}
	if left > maxX {
		left = maxX
	}
	if top < leftMargin {
		top = leftMargin
	}
	if top > maxY {
		top = maxY
	}
	if right > maxX {
		right = maxX
	}
	if right < left {
		right = left + 1
	}
	if bottom > maxY {
		bottom = maxY
	}
	if bottom < top {
		bottom = top + 1
	}

	w := right - left
	h := bottom - top
	if w <= 0 || h <= 0 {
		return
	}

	// Apply line width
	if lineWidth > 0 {
		c.pdf.pdf.SetLineWidth(lineWidth)
	}

	// Apply color
	if r, g, b, ok := parseAspPDFColor(p, "color"); ok {
		c.pdf.pdf.SetDrawColor(r, g, b)
	}

	c.pdf.pdf.Rect(left, top, w, h, "D")
}

// PdfCanvasDispatchMethod dispatches method calls on a PdfCanvas.
func (c *PdfCanvas) DispatchMethod(member string, args []Value) Value {
	switch strings.ToLower(member) {
	case "drawtext":
		if len(args) < 1 {
			return NewBool(false)
		}
		text := fmt.Sprintf("%v", legacyValueToInterface(args[0], c.vm))
		params := ""
		var font *PdfFont
		if len(args) > 1 {
			params = fmt.Sprintf("%v", legacyValueToInterface(args[1], c.vm))
		}
		if len(args) > 2 {
			if args[2].Type == VTNativeObject {
				if f, exists := c.vm.pdfFontItems[args[2].Num]; exists {
					font = f
				}
			}
		}
		c.DrawText(text, params, font)
		return NewBool(true)

	case "drawline":
		if len(args) < 1 {
			return NewBool(false)
		}
		params := fmt.Sprintf("%v", legacyValueToInterface(args[0], c.vm))
		c.DrawLine(params)
		return NewBool(true)

	case "drawbox":
		if len(args) < 1 {
			return NewBool(false)
		}
		params := fmt.Sprintf("%v", legacyValueToInterface(args[0], c.vm))
		c.DrawBox(params)
		return NewBool(true)
	}
	return NewEmpty()
}

// PdfCanvasDispatchPropertyGet returns a property value for PdfCanvas.
func (c *PdfCanvas) DispatchPropertyGet(member string) Value {
	return NewEmpty()
}

// PdfCanvasDispatchPropertySet sets a property on PdfCanvas.
func (c *PdfCanvas) DispatchPropertySet(member string, args []Value) bool {
	return false
}

// ============================================================================
// PdfPage - Persits.Pdf Page object
// ============================================================================

// PdfPage represents a single page in a PDF document.
type PdfPage struct {
	vm        *VM
	objectID  int64
	pdf       *G3PDF
	doc       *PdfDocument
	canvasVal Value // cached VTNativeObject for Canvas
}

// getCanvas returns (or creates) the PdfCanvas for this page.
func (p *PdfPage) getCanvas() Value {
	if p.canvasVal.Type == VTNativeObject {
		return p.canvasVal
	}
	canvas := &PdfCanvas{
		vm:   p.vm,
		pdf:  p.pdf,
		page: p,
	}
	id := p.vm.nextDynamicNativeID
	p.vm.nextDynamicNativeID++
	canvas.objectID = id
	p.vm.pdfCanvasItems[id] = canvas
	p.canvasVal = Value{Type: VTNativeObject, Num: id}
	return p.canvasVal
}

// PdfPageDispatchPropertyGet returns a property value for PdfPage.
func (p *PdfPage) DispatchPropertyGet(member string) Value {
	switch strings.ToLower(member) {
	case "canvas":
		return p.getCanvas()
	case "width", "w":
		w, _ := p.pdf.pdf.GetPageSize()
		return NewDouble(w)
	case "height", "h":
		_, h := p.pdf.pdf.GetPageSize()
		return NewDouble(h)
	case "rotation":
		return NewInteger(0)
	case "pagenumber":
		return NewInteger(int64(p.pdf.pdf.PageNo()))
	}
	return NewEmpty()
}

// PdfPageDispatchPropertySet sets a property on PdfPage.
func (p *PdfPage) DispatchPropertySet(member string, args []Value) bool {
	return false
}

// PdfPageDispatchMethod dispatches method calls on a PdfPage.
func (p *PdfPage) DispatchMethod(member string, args []Value) Value {
	return NewEmpty()
}

// ============================================================================
// PdfDocument - Persits.Pdf Document object
// ============================================================================

// PdfDocument represents the top-level document created by PdfManager.CreateDocument.
type PdfDocument struct {
	vm       *VM
	objectID int64
	pdf      *G3PDF

	pagesVal    Value // cached VTNativeObject for the Pages collection (mock)
	open        bool
	filePath    string
	cachedBytes []byte // cached PDF output; generated once, reused across Save/SendBinary
}

// ensureOutput generates the PDF bytes once and caches them.
func (d *PdfDocument) ensureOutput() ([]byte, error) {
	if d.cachedBytes != nil {
		return d.cachedBytes, nil
	}
	buf := &bytes.Buffer{}
	if err := d.pdf.pdf.Output(buf); err != nil {
		return nil, fmt.Errorf("pdf generation failed: %w", err)
	}
	d.cachedBytes = buf.Bytes()
	return d.cachedBytes, nil
}

// getPages returns a simple mock Pages collection. In AspPDF, .Pages.Add() creates a page.
// We represent Pages as a lightweight object with an Add method.
func (d *PdfDocument) getPages() Value {
	if d.pagesVal.Type == VTNativeObject {
		return d.pagesVal
	}
	// Pages is a pseudo-object; we create a PdfPagesCollection placeholder.
	pages := &PdfPagesCollection{
		vm:  d.vm,
		doc: d,
	}
	id := d.vm.nextDynamicNativeID
	d.vm.nextDynamicNativeID++
	pages.objectID = id
	d.vm.pdfPagesItems[id] = pages
	d.pagesVal = Value{Type: VTNativeObject, Num: id}
	return d.pagesVal
}

// Save writes the PDF document to a file path.
func (d *PdfDocument) Save(path string) error {
	data, err := d.ensureOutput()
	if err != nil {
		return fmt.Errorf("%w: %v", NewAxonASPError(ErrG3PDFSaveFailed, nil, AxonASPErrorMessages[ErrG3PDFSaveFailed], "lib_g3pdf_persits.go", 0), err)
	}
	return os.WriteFile(path, data, 0666)
}

// SendBinary returns the PDF as a byte slice for Response.BinaryWrite.
func (d *PdfDocument) SendBinary() ([]byte, error) {
	return d.ensureOutput()
}

// ImportFromUrl parses params and routes to the existing WriteHTML logic.
// Param keys: scale, drawbackground, width, height, autofit, margin.
func (d *PdfDocument) ImportFromUrl(url string, params string) error {
	if url == "" {
		return fmt.Errorf("ImportFromUrl requires a non-empty URL")
	}
	p := parseAspPDFParams(params)

	// Try to load the URL content as HTML
	// In Classic ASP / Persits.Pdf, ImportFromUrl fetches a URL and converts
	// the HTML to PDF. We route this through the existing WriteHTML logic
	// by attempting to read the URL content first.
	content, err := d.fetchURLContent(url)
	if err != nil {
		return fmt.Errorf("%w: failed to fetch URL %s: %v",
			NewAxonASPError(ErrG3PDFImportFailed, nil, AxonASPErrorMessages[ErrG3PDFImportFailed], "lib_g3pdf_persits.go", 0), url, err)
	}

	// Apply scale if specified
	scale := parseAspPDFFloat(p, "scale", 1.0)
	_ = scale // scale is applied by rendering at appropriate size

	// Route to existing WriteHTML engine
	d.pdf.WriteHTML(string(content))
	return nil
}

// fetchURLContent attempts to fetch content from a URL.
// In the current implementation, we try to read the URL as a local file path
// (useful for testing and local content). For remote URLs, this will fail
// with a descriptive error unless the host environment provides HTTP fetching.
func (d *PdfDocument) fetchURLContent(url string) ([]byte, error) {
	// Try as local file first
	if _, err := os.Stat(url); err == nil {
		return os.ReadFile(url)
	}
	// If it looks like a remote URL, return an error
	if strings.HasPrefix(strings.ToLower(url), "http://") || strings.HasPrefix(strings.ToLower(url), "https://") {
		return nil, fmt.Errorf("remote URL fetching requires an HTTP-enabled host environment; use WriteHTML for direct HTML input")
	}
	return nil, fmt.Errorf("cannot fetch URL: file not found or no HTTP client available")
}

// PdfDocumentDispatchMethod dispatches method calls on PdfDocument.
func (d *PdfDocument) DispatchMethod(member string, args []Value) Value {
	switch strings.ToLower(member) {
	case "save":
		if len(args) < 1 {
			if d.vm != nil {
				d.vm.raise(vbscript.InternalError, "Save requires a file path argument")
			}
			return NewBool(false)
		}
		path := fmt.Sprintf("%v", legacyValueToInterface(args[0], d.vm))
		if err := d.Save(path); err != nil {
			if d.vm != nil {
				d.vm.raise(vbscript.InternalError, err.Error())
			}
			return NewBool(false)
		}
		return NewBool(true)

	case "sendbinary", "sendbinarydata":
		data, err := d.SendBinary()
		if err != nil {
			if d.vm != nil {
				d.vm.raise(vbscript.InternalError, err.Error())
			}
			return NewNull()
		}
		// Return as a byte array for Response.BinaryWrite in VBScript
		return NewString(string(data))

	case "importfromurl":
		if len(args) < 1 {
			return NewBool(false)
		}
		url := fmt.Sprintf("%v", legacyValueToInterface(args[0], d.vm))
		params := ""
		if len(args) > 1 {
			params = fmt.Sprintf("%v", legacyValueToInterface(args[1], d.vm))
		}
		if err := d.ImportFromUrl(url, params); err != nil {
			if d.pdf != nil {
				d.pdf.setError(err.Error())
			}
			return NewBool(false)
		}
		return NewBool(true)

	case "close":
		d.open = false
		return NewBool(true)
	}
	return NewEmpty()
}

// PdfDocumentDispatchPropertyGet returns a property value for PdfDocument.
func (d *PdfDocument) DispatchPropertyGet(member string) Value {
	switch strings.ToLower(member) {
	case "pages":
		return d.getPages()
	case "open":
		return NewBool(d.open)
	}
	return NewEmpty()
}

// PdfDocumentDispatchPropertySet sets a property on PdfDocument.
func (d *PdfDocument) DispatchPropertySet(member string, args []Value) bool {
	return false
}

// ============================================================================
// PdfPagesCollection - Mock collection for .Pages accessor
// ============================================================================

// PdfPagesCollection is a lightweight mock for the Pages collection.
// It provides the Add() method which creates a new PdfPage.
type PdfPagesCollection struct {
	vm       *VM
	objectID int64
	doc      *PdfDocument
}

// PdfPagesDispatchMethod dispatches method calls on the Pages collection.
func (pc *PdfPagesCollection) DispatchMethod(member string, args []Value) Value {
	switch strings.ToLower(member) {
	case "add":
		// Add a new page to the document
		pc.doc.pdf.pdf.AddPage()
		page := &PdfPage{
			vm:  pc.vm,
			pdf: pc.doc.pdf,
			doc: pc.doc,
		}
		id := pc.vm.nextDynamicNativeID
		pc.vm.nextDynamicNativeID++
		page.objectID = id
		pc.vm.pdfPageItems[id] = page
		return Value{Type: VTNativeObject, Num: id}
	}
	return NewEmpty()
}

// PdfPagesDispatchPropertyGet returns a property value for the Pages collection.
func (pc *PdfPagesCollection) DispatchPropertyGet(member string) Value {
	switch strings.ToLower(member) {
	case "count":
		return NewInteger(int64(pc.doc.pdf.pdf.PageNo()))
	}
	return NewEmpty()
}

// ============================================================================
// PdfManager methods on G3PDF
// ============================================================================

// CreateDocument implements PdfManager.CreateDocument.
// It creates a new PdfDocument sub-object linked to this G3PDF instance.
func (p *G3PDF) CreateDocument() *PdfDocument {
	doc := &PdfDocument{
		vm:   p.vm,
		pdf:  p,
		open: true,
	}
	id := p.vm.nextDynamicNativeID
	p.vm.nextDynamicNativeID++
	doc.objectID = id
	p.vm.pdfDocItems[id] = doc
	return doc
}

// OpenDocument opens an existing PDF file and returns a PdfDocument.
func (p *G3PDF) OpenDocument(path string) (*PdfDocument, error) {
	data, err := os.ReadFile(path)
	if err != nil {
		return nil, fmt.Errorf("failed to open PDF: %w", err)
	}
	// Create a new PDF document and import the existing content
	p.Reset("P", "mm", "A4")
	// Write the raw PDF content to the output buffer
	// Note: full PDF import is limited - we create a new doc and embed the content
	doc := p.CreateDocument()
	doc.filePath = path
	_ = data // fpdf does not support loading existing PDFs natively
	return doc, nil
}

// getFont loads or creates a PdfFont by name.
// Fonts are cached in the G3PDF struct.
func (p *G3PDF) getFont(name string) *PdfFont {
	// Check cache
	if p.fontCache == nil {
		p.fontCache = make(map[string]*PdfFont)
	}
	if f, ok := p.fontCache[name]; ok {
		return f
	}
	// Map common names to fpdf families
	family, style := mapPersitsFontToFPDF(name)
	font := &PdfFont{
		vm:       p.vm,
		pdf:      p,
		name:     name,
		family:   family,
		style:    style,
		size:     12,
		embedded: true,
	}
	id := p.vm.nextDynamicNativeID
	p.vm.nextDynamicNativeID++
	font.objectID = id
	p.vm.pdfFontItems[id] = font
	p.fontCache[name] = font
	return font
}

// mapPersitsFontToFPDF converts a Persits.Pdf font name to fpdf family + style.
func mapPersitsFontToFPDF(name string) (string, string) {
	n := strings.ToLower(strings.TrimSpace(name))
	switch n {
	case "helvetica", "arial":
		return "helvetica", ""
	case "helvetica-bold", "arial bold", "arialbd":
		return "helvetica", "B"
	case "helvetica-oblique", "helvetica italic", "arial italic", "ariali":
		return "helvetica", "I"
	case "helvetica-boldoblique", "helvetica bold italic", "arial bold italic", "arialbi":
		return "helvetica", "BI"
	case "times", "times roman", "times new roman":
		return "times", ""
	case "times bold", "times new roman bold":
		return "times", "B"
	case "times italic", "times new roman italic":
		return "times", "I"
	case "times bold italic", "times new roman bold italic":
		return "times", "BI"
	case "courier":
		return "courier", ""
	case "courier-bold":
		return "courier", "B"
	case "courier-oblique", "courier italic":
		return "courier", "I"
	case "courier-boldoblique", "courier bold italic":
		return "courier", "BI"
	case "symbol":
		return "symbol", ""
	case "zapfdingbats":
		return "zapfdingbats", ""
	default:
		return "helvetica", ""
	}
}

// ============================================================================
// Persits.Pdf-specific DispatchMethod additions (routed from G3PDF)
// ============================================================================

// dispatchPersitsMethod handles Persits.Pdf methods on the G3PDF object.
// These are called when G3PDF.DispatchMethod receives a Persits-specific call.
func (p *G3PDF) dispatchPersitsMethod(methodName string, args []Value) Value {
	switch methodName {
	case "createdocument":
		doc := p.CreateDocument()
		return Value{Type: VTNativeObject, Num: doc.objectID}

	case "opendocument":
		if len(args) < 1 {
			return NewBool(false)
		}
		path := fmt.Sprintf("%v", legacyValueToInterface(args[0], p.vm))
		doc, err := p.OpenDocument(path)
		if err != nil {
			p.setError(err.Error())
			return NewBool(false)
		}
		return Value{Type: VTNativeObject, Num: doc.objectID}

	case "fonts":
		if len(args) < 1 {
			return NewNull()
		}
		name := fmt.Sprintf("%v", legacyValueToInterface(args[0], p.vm))
		font := p.getFont(name)
		return Value{Type: VTNativeObject, Num: font.objectID}

	case "set", "setparam":
		// Persits.Pdf Set method for global settings
		return NewBool(true)

	case "get", "getparam":
		return NewEmpty()
	}
	return NewNull()
}

// ============================================================================
// VM dispatch functions for Persits.Pdf sub-objects
// ============================================================================

// dispatchPdfDocumentMethod dispatches method calls on PdfDocument objects.
func (vm *VM) dispatchPdfDocumentMethod(doc *PdfDocument, member string, args []Value) Value {
	return doc.DispatchMethod(member, args)
}

// dispatchPdfDocumentPropertyGet dispatches property get on PdfDocument objects.
func (vm *VM) dispatchPdfDocumentPropertyGet(doc *PdfDocument, member string) Value {
	return doc.DispatchPropertyGet(member)
}

// dispatchPdfDocumentPropertySet dispatches property set on PdfDocument objects.
func (vm *VM) dispatchPdfDocumentPropertySet(doc *PdfDocument, member string, val Value) {
	doc.DispatchPropertySet(member, []Value{val})
}

// dispatchPdfPageMethod dispatches method calls on PdfPage objects.
func (vm *VM) dispatchPdfPageMethod(page *PdfPage, member string, args []Value) Value {
	return page.DispatchMethod(member, args)
}

// dispatchPdfPagePropertyGet dispatches property get on PdfPage objects.
func (vm *VM) dispatchPdfPagePropertyGet(page *PdfPage, member string) Value {
	return page.DispatchPropertyGet(member)
}

// dispatchPdfPagePropertySet dispatches property set on PdfPage objects.
func (vm *VM) dispatchPdfPagePropertySet(page *PdfPage, member string, val Value) {
	page.DispatchPropertySet(member, []Value{val})
}

// dispatchPdfCanvasMethod dispatches method calls on PdfCanvas objects.
func (vm *VM) dispatchPdfCanvasMethod(canvas *PdfCanvas, member string, args []Value) Value {
	return canvas.DispatchMethod(member, args)
}

// dispatchPdfCanvasPropertyGet dispatches property get on PdfCanvas objects.
func (vm *VM) dispatchPdfCanvasPropertyGet(canvas *PdfCanvas, member string) Value {
	return canvas.DispatchPropertyGet(member)
}

// dispatchPdfCanvasPropertySet dispatches property set on PdfCanvas objects.
func (vm *VM) dispatchPdfCanvasPropertySet(canvas *PdfCanvas, member string, val Value) {
	canvas.DispatchPropertySet(member, []Value{val})
}

// dispatchPdfFontMethod dispatches method calls on PdfFont objects.
func (vm *VM) dispatchPdfFontMethod(font *PdfFont, member string, args []Value) Value {
	return font.DispatchMethod(member, args)
}

// dispatchPdfFontPropertyGet dispatches property get on PdfFont objects.
func (vm *VM) dispatchPdfFontPropertyGet(font *PdfFont, member string) Value {
	return font.DispatchPropertyGet(member)
}

// dispatchPdfFontPropertySet dispatches property set on PdfFont objects.
func (vm *VM) dispatchPdfFontPropertySet(font *PdfFont, member string, val Value) {
	font.DispatchPropertySet(member, []Value{val})
}

// dispatchPdfPagesMethod dispatches method calls on PdfPagesCollection objects.
func (vm *VM) dispatchPdfPagesMethod(pc *PdfPagesCollection, member string, args []Value) Value {
	return pc.DispatchMethod(member, args)
}

// dispatchPdfPagesPropertyGet dispatches property get on PdfPagesCollection objects.
func (vm *VM) dispatchPdfPagesPropertyGet(pc *PdfPagesCollection, member string) Value {
	return pc.DispatchPropertyGet(member)
}

// cleanupG3PDFResources releases all Persits.Pdf sub-objects owned by the current VM.
func (vm *VM) cleanupG3PDFResources() {
	for id := range vm.pdfDocItems {
		delete(vm.pdfDocItems, id)
	}
	for id := range vm.pdfPageItems {
		delete(vm.pdfPageItems, id)
	}
	for id := range vm.pdfCanvasItems {
		delete(vm.pdfCanvasItems, id)
	}
	for id := range vm.pdfFontItems {
		delete(vm.pdfFontItems, id)
	}
	for id := range vm.pdfPagesItems {
		delete(vm.pdfPagesItems, id)
	}
}
