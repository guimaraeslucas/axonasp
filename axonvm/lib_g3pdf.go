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
 * Attribution Notice:
 * If this software is used in other projects, the name "AxonASP Server"
 * must be cited in the documentation or "About" section.
 *
 * Contribution Policy:
 * Modifications to the core source code of AxonASP Server must be
 * made available under this same license terms.
 *
 * Third-Party Notice:
 * G3PDF is a wrapper around go-pdf/fpdf library
 * Original FPDF authored by Olivier Plathey
 */
package axonvm

import (
	"bytes"
	"fmt"
	"math"
	"os"
	"strconv"
	"strings"
	"time"

	"codeberg.org/go-pdf/fpdf"
)

// G3PDF is a high-performance AxonASP wrapper around go-pdf/fpdf
// designed for memory efficiency and direct bytecode execution
type G3PDF struct {
	vm *VM

	// Core PDF engine from go-pdf/fpdf
	pdf *fpdf.Fpdf

	// State tracking for compatibility
	lastError string
	fontpath  string

	// Metadata storage
	stdPageSizes map[string][2]float64
	title        string
	author       string
	subject      string
	keywords     string
	creator      string
	compression  bool
	aliasNbPages string

	// HTML rendering state (used during WriteHTML execution only)
	htmlState *g3PDFHTMLState
}

// g3PDFHTMLState manages state during HTML parsing
type g3PDFHTMLState struct {
	pdf *G3PDF

	// Tag stack for nested elements
	tagStack []string

	// Style stack for nested tags
	styles []*g3PDFHTMLStyle

	// Text formatting state (use booleans for simplicity)
	bold      bool
	italic    bool
	underline bool

	// Bold/italic/underline counters for proper nesting
	boldCount      int
	italicCount    int
	underlineCount int

	// Current font size
	fontSize float64

	// Text color RGB (0-255)
	textColor [3]int

	// Link and URL handling
	href string

	// Preformatted text mode
	pre bool

	// List handling
	listDepth int
	listType  string // "O" for ordered, "U" for unordered
	listCount int
	listStack []*g3PDFHTMLListState

	// Table state
	inTable        bool
	inRow          bool
	tableBorder    int
	rowStartY      float64
	maxRowHeight   float64
	tableColWidths map[int]float64
	cellPadding    float64
	cellSpacing    float64
	colIndex       int

	// Cell state
	cellText    string
	tdBegin     bool
	thBegin     bool
	tdWidth     float64
	tdHeight    float64
	tdAlign     string
	tdWidthAttr string
	tdBgColor   bool
	trBgColor   bool
	tdColorR    float64
	tdColorG    float64
	tdColorB    float64
	tdColorSet  bool

	// Alignment
	currAlign string

	// Script positioning
	scriptActive    bool
	scriptDeltaY    float64
	defaultFontSize float64

	// Color management
	colorSet bool
	fontSet  bool
}

// g3PDFHTMLStyle tracks CSS/style state for proper restoration
type g3PDFHTMLStyle struct {
	fontFamily string
	fontStyle  string
	fontSize   float64
	textColorR float64
	textColorG float64
	textColorB float64
	colorSet   bool
	bold       bool
	italic     bool
	underline  bool
	href       string
}

// g3PDFHTMLListState preserves list nesting state
type g3PDFHTMLListState struct {
	listType  string
	listCount int
}

// NewG3PDF creates a new G3PDF instance optimized for AxonASP
func NewG3PDF(ctx *VM) *G3PDF {
	p := &G3PDF{
		vm:           ctx,
		pdf:          fpdf.New("P", "mm", "A4", ""),
		compression:  true,
		aliasNbPages: "{nb}",
	}

	// Initialize standard page sizes
	p.stdPageSizes = map[string][2]float64{
		"a3":     {841.89 / 2.834645669, 1190.55 / 2.834645669},
		"a4":     {595.28 / 2.834645669, 841.89 / 2.834645669},
		"a5":     {420.94 / 2.834645669, 595.28 / 2.834645669},
		"letter": {612.0 / 2.834645669, 792.0 / 2.834645669},
		"legal":  {612.0 / 2.834645669, 1008.0 / 2.834645669},
	}

	// Apply default settings
	p.pdf.SetCompression(p.compression)
	p.pdf.SetCreator("G3Pix AxonASP", false)
	p.pdf.SetCreationDate(time.Now())

	return p
}

// Reset reinitializes the PDF document with new settings
func (p *G3PDF) Reset(orientation, unit, size string) {
	// Validate inputs
	orientation = strings.ToUpper(strings.TrimSpace(orientation))
	if orientation != "P" && orientation != "L" {
		orientation = "P"
	}

	unit = strings.ToLower(strings.TrimSpace(unit))
	if unit == "" {
		unit = "mm"
	}

	size = strings.ToUpper(strings.TrimSpace(size))
	if size == "" {
		size = "A4"
	}

	// Create new PDF instance
	p.pdf = fpdf.New(orientation, unit, size, "")
	p.pdf.SetCompression(p.compression)

	// Restore metadata if set
	if p.title != "" {
		p.pdf.SetTitle(p.title, true)
	}
	if p.author != "" {
		p.pdf.SetAuthor(p.author, true)
	}
	if p.subject != "" {
		p.pdf.SetSubject(p.subject, true)
	}
	if p.keywords != "" {
		p.pdf.SetKeywords(p.keywords, true)
	}
	if p.creator != "" {
		p.pdf.SetCreator(p.creator+" (G3Pix AxonASP)", true)
	} else {
		p.pdf.SetCreator("G3Pix AxonASP", false)
	}
	p.lastError = ""
}

// AddPage adds a new page to the document
// orientation: empty string keeps current, "P" for portrait, "L" for landscape
func (p *G3PDF) AddPage(orientation, size string, rotation int) {
	orientation = strings.ToUpper(strings.TrimSpace(orientation))
	if size == "" && orientation == "" {
		p.pdf.AddPage()
		return
	}

	// Use AddPageFormat for custom orientation or size
	pageSize := fpdf.SizeType{}
	if size != "" {
		pageSize = p.pdf.GetPageSizeStr(strings.ToUpper(strings.TrimSpace(size)))
	}
	if orientation == "" {
		orientation = "P"
	}
	p.pdf.AddPageFormat(orientation, pageSize)
}

// Close closes the PDF document
func (p *G3PDF) Close() {
	p.lastError = ""
}

// SetFont sets the current font for text rendering
func (p *G3PDF) SetFont(family, style string, size float64) {
	if family == "" {
		family = "helvetica"
	}
	if style == "" {
		style = ""
	}
	if size <= 0 {
		size = 12
	}

	family = strings.ToLower(strings.TrimSpace(family))
	style = strings.ToUpper(strings.TrimSpace(style))

	p.pdf.SetFont(family, style, size)
}

// SetFontSize sets the font size only
func (p *G3PDF) SetFontSize(size float64) {
	if size <= 0 {
		size = 12
	}
	p.pdf.SetFontSize(size)
}

// setError sets the internal error state
func (p *G3PDF) setError(msg string) {
	p.lastError = msg
}

// SetTextColor sets the text color using RGB components (0-255 scale)
// If g and b are NaN, treats r as grayscale value
func (p *G3PDF) SetTextColor(r, g, b float64) {
	if math.IsNaN(g) || math.IsNaN(b) {
		// Grayscale mode - scale 0-255 to 0-100
		gray := int(math.Max(0, math.Min(255, r))) / 255 * 100
		p.pdf.SetTextColor(gray, 0, 0)
	} else {
		// RGB mode - convert 0-255 to 0-255 for fpdf
		rVal := int(math.Max(0, math.Min(255, r)))
		gVal := int(math.Max(0, math.Min(255, g)))
		bVal := int(math.Max(0, math.Min(255, b)))
		p.pdf.SetTextColor(rVal, gVal, bVal)
	}
}

// SetDrawColor sets the drawing color for lines and borders
func (p *G3PDF) SetDrawColor(r, g, b float64) {
	if math.IsNaN(g) || math.IsNaN(b) {
		// Grayscale mode
		gray := int(math.Max(0, math.Min(255, r))) / 255 * 100
		p.pdf.SetDrawColor(gray, 0, 0)
	} else {
		// RGB mode
		rVal := int(math.Max(0, math.Min(255, r)))
		gVal := int(math.Max(0, math.Min(255, g)))
		bVal := int(math.Max(0, math.Min(255, b)))
		p.pdf.SetDrawColor(rVal, gVal, bVal)
	}
}

// SetFillColor sets the fill color for cell backgrounds
func (p *G3PDF) SetFillColor(r, g, b float64) {
	if math.IsNaN(g) || math.IsNaN(b) {
		// Grayscale mode
		gray := int(math.Max(0, math.Min(255, r))) / 255 * 100
		p.pdf.SetFillColor(gray, 0, 0)
	} else {
		// RGB mode
		rVal := int(math.Max(0, math.Min(255, r)))
		gVal := int(math.Max(0, math.Min(255, g)))
		bVal := int(math.Max(0, math.Min(255, b)))
		p.pdf.SetFillColor(rVal, gVal, bVal)
	}
}

// SetLineWidth sets the thickness of drawn lines and borders
func (p *G3PDF) SetLineWidth(width float64) {
	if width <= 0 {
		width = 0.1
	}
	p.pdf.SetLineWidth(width)
}

// SetMargins sets left, top, and right page margins
// right is optional; if nil, defaults to same as left
func (p *G3PDF) SetMargins(left, top float64, right *float64) {
	if left <= 0 {
		left = 10.0
	}
	if top <= 0 {
		top = 10.0
	}
	r := left // default: same as left
	if right != nil && *right > 0 {
		r = *right
	}
	p.pdf.SetMargins(left, top, r)
}

// SetLeftMargin sets the left margin
func (p *G3PDF) SetLeftMargin(margin float64) {
	if margin > 0 {
		p.pdf.SetLeftMargin(margin)
	}
}

// SetTopMargin sets the top margin
func (p *G3PDF) SetTopMargin(margin float64) {
	if margin > 0 {
		p.pdf.SetTopMargin(margin)
	}
}

// SetRightMargin sets the right margin
func (p *G3PDF) SetRightMargin(margin float64) {
	if margin > 0 {
		p.pdf.SetRightMargin(margin)
	}
}

// SetAutoPageBreak enables automatic page breaks
func (p *G3PDF) SetAutoPageBreak(auto bool, margin float64) {
	if margin <= 0 {
		margin = 0
	}
	p.pdf.SetAutoPageBreak(auto, margin)
}

// SetX sets the X coordinate of the current position
func (p *G3PDF) SetX(x float64) {
	p.pdf.SetX(x)
}

// SetY sets the Y coordinate and resets X to left margin
func (p *G3PDF) SetY(y float64, _ bool) {
	p.pdf.SetY(y)
}

// SetXY sets both X and Y coordinates in one call
func (p *G3PDF) SetXY(x, y float64) {
	p.pdf.SetXY(x, y)
}

// Ln moves to next line with optional offset
func (p *G3PDF) Ln(offset float64) {
	if offset < 0 {
		offset = 0
	}
	p.pdf.Ln(offset)
}

// Line draws a line from (x1,y1) to (x2,y2) using current draw color and line width
func (p *G3PDF) Line(x1, y1, x2, y2 float64) {
	p.pdf.Line(x1, y1, x2, y2)
}

// Rect draws a rectangle at (x,y) with specified width and height
// style: empty or "S" = draw border only, "F" = fill, "DF" or "FD" = both
func (p *G3PDF) Rect(x, y, w, h float64, style string) {
	style = strings.ToUpper(strings.TrimSpace(style))
	if style == "" {
		style = "D"
	}
	p.pdf.Rect(x, y, w, h, style)
}

// Cell prints a rectangular cell with text at the current position
func (p *G3PDF) Cell(w, h float64, txt string, border interface{}, ln int, align string, fill bool, link interface{}) {
	// Convert border parameter (can be 0, 1, string like "LTRB")
	var borderStr string
	if b, ok := border.(string); ok {
		borderStr = strings.ToUpper(b)
	} else {
		if toBool(border) {
			borderStr = "1"
		}
	}

	// Normalize alignment
	align = strings.ToUpper(strings.TrimSpace(align))

	// Handle link parameter - can be internal link ID (int) or external URL (string)
	linkID := 0
	linkStr := ""
	if l, ok := link.(int); ok {
		linkID = l
	} else if l, ok := link.(int64); ok {
		linkID = int(l)
	} else if l, ok := link.(string); ok && l != "" {
		linkStr = l
	}

	p.pdf.CellFormat(w, h, txt, borderStr, ln, align, fill, linkID, linkStr)
}

// MultiCell prints text with multiple lines
func (p *G3PDF) MultiCell(w, h float64, txt string, border interface{}, align string, fill bool) {
	// Convert border parameter
	var borderStr string
	if b, ok := border.(string); ok {
		borderStr = strings.ToUpper(b)
	} else {
		if toBool(border) {
			borderStr = "1"
		}
	}

	// Normalize alignment
	align = strings.ToUpper(strings.TrimSpace(align))

	p.pdf.MultiCell(w, h, txt, borderStr, align, fill)
}

// Write prints flowing text starting at current position
// link can be an integer link ID (from AddLink) or a URL string
func (p *G3PDF) Write(h float64, txt string, link interface{}) {
	switch l := link.(type) {
	case int:
		if l > 0 {
			p.pdf.WriteLinkID(h, txt, l)
			return
		}
	case int64:
		if l > 0 {
			p.pdf.WriteLinkID(h, txt, int(l))
			return
		}
	case string:
		if l != "" {
			p.pdf.WriteLinkString(h, txt, l)
			return
		}
	}
	p.pdf.Write(h, txt)
}

// Text prints text at absolute coordinates (x, y)
func (p *G3PDF) Text(x, y float64, txt string) {
	p.pdf.Text(x, y, txt)
}

// Image inserts an image at specified position
// typ can be "JPG", "PNG", "GIF", "BMP" etc
func (p *G3PDF) Image(path string, x, y, w, h float64, typ string, link interface{}) {
	path = strings.TrimSpace(path)
	if path == "" {
		return
	}

	linkID := 0
	if l, ok := link.(int); ok {
		linkID = l
	} else if l, ok := link.(int64); ok {
		linkID = int(l)
	}

	// If x or y is NaN, use current position
	if math.IsNaN(x) {
		x = p.pdf.GetX()
	}
	if math.IsNaN(y) {
		y = p.pdf.GetY()
	}

	typ = strings.ToUpper(strings.TrimSpace(typ))

	// Try to insert image - fpdf will handle file resolution
	opts := fpdf.ImageOptions{
		ImageType: typ,
	}

	p.pdf.ImageOptions(path, x, y, w, h, false, opts, linkID, "")
	if err := p.pdf.Error(); err != nil {
		p.setError(fmt.Sprintf("image error: %v", err))
	}
}

// SetTitle sets the PDF title metadata
func (p *G3PDF) SetTitle(text string, isUTF8 bool) {
	p.title = text
	p.pdf.SetTitle(text, isUTF8)
}

// SetAuthor sets the PDF author metadata
func (p *G3PDF) SetAuthor(text string, isUTF8 bool) {
	p.author = text
	p.pdf.SetAuthor(text, isUTF8)
}

// SetSubject sets the PDF subject metadata
func (p *G3PDF) SetSubject(text string, isUTF8 bool) {
	p.subject = text
	p.pdf.SetSubject(text, isUTF8)
}

// SetKeywords sets the PDF keywords metadata
func (p *G3PDF) SetKeywords(text string, isUTF8 bool) {
	p.keywords = text
	p.pdf.SetKeywords(text, isUTF8)
}

// SetCreator sets the PDF creator metadata
func (p *G3PDF) SetCreator(text string, isUTF8 bool) {
	p.creator = text
	creatorStr := "AxonASP Library"
	if text != "" {
		creatorStr = text + " (AxonASP)"
	}
	p.pdf.SetCreator(creatorStr, isUTF8)
}

// AliasNbPages sets the alias for the total number of pages
func (p *G3PDF) AliasNbPages(alias string) {
	if alias == "" {
		alias = "{nb}"
	}
	p.aliasNbPages = alias
	p.pdf.AliasNbPages(alias)
}

// SetDisplayMode sets the display mode for PDF viewers
func (p *G3PDF) SetDisplayMode(zoom, layout string) {
	zoom = strings.ToLower(strings.TrimSpace(zoom))
	layout = strings.ToLower(strings.TrimSpace(layout))

	p.pdf.SetDisplayMode(zoom, layout)
}

// SetCompression enables or disables PDF compression
func (p *G3PDF) SetCompression(enable bool) {
	p.compression = enable
	p.pdf.SetCompression(enable)
}

// GetStringWidth returns the width of a text string in current font
func (p *G3PDF) GetStringWidth(txt string) float64 {
	return p.pdf.GetStringWidth(txt)
}

// Need to be implemented in our current library, to when the user use in HTML this colors, it gets automaticaly converted to RGB
func htmlColorToRGB(color string) (int, int, int, bool) {
	if color == "" {
		return 0, 0, 0, false
	}
	color = strings.TrimSpace(strings.ToUpper(color))
	basic := map[string]string{
		"ALICEBLUE":            "#F0F8FF",
		"ANTIQUEWHITE":         "#FAEBD7",
		"AQUA":                 "#00FFFF",
		"AQUAMARINE":           "#7FFFD4",
		"AZURE":                "#F0FFFF",
		"BEIGE":                "#F5F5DC",
		"BISQUE":               "#FFE4C4",
		"BLACK":                "#000000",
		"BLANCHEDALMOND":       "#FFEBCD",
		"BLUE":                 "#0000FF",
		"BLUEVIOLET":           "#8A2BE2",
		"BROWN":                "#A52A2A",
		"BURLYWOOD":            "#DEB887",
		"CADETBLUE":            "#5F9EA0",
		"CHARTREUSE":           "#7FFF00",
		"CHOCOLATE":            "#D2691E",
		"CORAL":                "#FF7F50",
		"CORNFLOWERBLUE":       "#6495ED",
		"CORNSILK":             "#FFF8DC",
		"CRIMSON":              "#DC143C",
		"CYAN":                 "#00FFFF",
		"DARKBLUE":             "#00008B",
		"DARKCYAN":             "#008B8B",
		"DARKGOLDENROD":        "#B8860B",
		"DARKGRAY":             "#A9A9A9",
		"DARKGREY":             "#A9A9A9",
		"DARKGREEN":            "#006400",
		"DARKKHAKI":            "#BDB76B",
		"DARKMAGENTA":          "#8B008B",
		"DARKOLIVEGREEN":       "#556B2F",
		"DARKORANGE":           "#FF8C00",
		"DARKORCHID":           "#9932CC",
		"DARKRED":              "#8B0000",
		"DARKSALMON":           "#E9967A",
		"DARKSEAGREEN":         "#8FBC8F",
		"DARKSLATEBLUE":        "#483D8B",
		"DARKSLATEGRAY":        "#2F4F4F",
		"DARKSLATEGREY":        "#2F4F4F",
		"DARKTURQUOISE":        "#00CED1",
		"DARKVIOLET":           "#9400D3",
		"DEEPPINK":             "#FF1493",
		"DEEPSKYBLUE":          "#00BFFF",
		"DIMGRAY":              "#696969",
		"DIMGREY":              "#696969",
		"DODGERBLUE":           "#1E90FF",
		"FIREBRICK":            "#B22222",
		"FLORALWHITE":          "#FFFAF0",
		"FORESTGREEN":          "#228B22",
		"FUCHSIA":              "#FF00FF",
		"GAINSBORO":            "#DCDCDC",
		"GHOSTWHITE":           "#F8F8FF",
		"GOLD":                 "#FFD700",
		"GOLDENROD":            "#DAA520",
		"GRAY":                 "#808080",
		"GREY":                 "#808080",
		"GREEN":                "#008000",
		"GREENYELLOW":          "#ADFF2F",
		"HONEYDEW":             "#F0FFF0",
		"HOTPINK":              "#FF69B4",
		"INDIANRED":            "#CD5C5C",
		"INDIGO":               "#4B0082",
		"IVORY":                "#FFFFF0",
		"KHAKI":                "#F0E68C",
		"LAVENDER":             "#E6E6FA",
		"LAVENDERBLUSH":        "#FFF0F5",
		"LAWNGREEN":            "#7CFC00",
		"LEMONCHIFFON":         "#FFFACD",
		"LIGHTBLUE":            "#ADD8E6",
		"LIGHTCORAL":           "#F08080",
		"LIGHTCYAN":            "#E0FFFF",
		"LIGHTGOLDENRODYELLOW": "#FAFAD2",
		"LIGHTGRAY":            "#D3D3D3",
		"LIGHTGREY":            "#D3D3D3",
		"LIGHTGREEN":           "#90EE90",
		"LIGHTPINK":            "#FFB6C1",
		"LIGHTSALMON":          "#FFA07A",
		"LIGHTSEAGREEN":        "#20B2AA",
		"LIGHTSKYBLUE":         "#87CEFA",
		"LIGHTSLATEGRAY":       "#778899",
		"LIGHTSLATEGREY":       "#778899",
		"LIGHTSTEELBLUE":       "#B0C4DE",
		"LIGHTYELLOW":          "#FFFFE0",
		"LIME":                 "#00FF00",
		"LIMEGREEN":            "#32CD32",
		"LINEN":                "#FAF0E6",
		"MAGENTA":              "#FF00FF",
		"MAROON":               "#800000",
		"MEDIUMAQUAMARINE":     "#66CDAA",
		"MEDIUMBLUE":           "#0000CD",
		"MEDIUMORCHID":         "#BA55D3",
		"MEDIUMPURPLE":         "#9370DB",
		"MEDIUMSEAGREEN":       "#3CB371",
		"MEDIUMSLATEBLUE":      "#7B68EE",
		"MEDIUMSPRINGGREEN":    "#00FA9A",
		"MEDIUMTURQUOISE":      "#48D1CC",
		"MEDIUMVIOLETRED":      "#C71585",
		"MIDNIGHTBLUE":         "#191970",
		"MINTCREAM":            "#F5FFFA",
		"MISTYROSE":            "#FFE4E1",
		"MOCCASIN":             "#FFE4B5",
		"NAVAJOWHITE":          "#FFDEAD",
		"NAVY":                 "#000080",
		"OLDLACE":              "#FDF5E6",
		"OLIVE":                "#808000",
		"OLIVEDRAB":            "#6B8E23",
		"ORANGE":               "#FFA500",
		"ORANGERED":            "#FF4500",
		"ORCHID":               "#DA70D6",
		"PALEGOLDENROD":        "#EEE8AA",
		"PALEGREEN":            "#98FB98",
		"PALETURQUOISE":        "#AFEEEE",
		"PALEVIOLETRED":        "#DB7093",
		"PAPAYAWHIP":           "#FFEFD5",
		"PEACHPUFF":            "#FFDAB9",
		"PERU":                 "#CD853F",
		"PINK":                 "#FFC0CB",
		"PLUM":                 "#DDA0DD",
		"POWDERBLUE":           "#B0E0E6",
		"PURPLE":               "#800080",
		"REBECCAPURPLE":        "#663399",
		"RED":                  "#FF0000",
		"ROSYBROWN":            "#BC8F8F",
		"ROYALBLUE":            "#4169E1",
		"SADDLEBROWN":          "#8B4513",
		"SALMON":               "#FA8072",
		"SANDYBROWN":           "#F4A460",
		"SEAGREEN":             "#2E8B57",
		"SEASHELL":             "#FFF5EE",
		"SIENNA":               "#A0522D",
		"SILVER":               "#C0C0C0",
		"SKYBLUE":              "#87CEEB",
		"SLATEBLUE":            "#6A5ACD",
		"SLATEGRAY":            "#708090",
		"SLATEGREY":            "#708090",
		"SNOW":                 "#FFFAFA",
		"SPRINGGREEN":          "#00FF7F",
		"STEELBLUE":            "#4682B4",
		"TAN":                  "#D2B48C",
		"TEAL":                 "#008080",
		"THISTLE":              "#D8BFD8",
		"TOMATO":               "#FF6347",
		"TURQUOISE":            "#40E0D0",
		"VIOLET":               "#EE82EE",
		"WHEAT":                "#F5DEB3",
		"WHITE":                "#FFFFFF",
		"WHITESMOKE":           "#F5F5F5",
		"YELLOW":               "#FFFF00",
		"YELLOWGREEN":          "#9ACD32",
	}
	if v, ok := basic[color]; ok {
		color = v
	}
	color = strings.TrimPrefix(color, "#")
	if len(color) == 3 {
		color = string([]byte{color[0], color[0], color[1], color[1], color[2], color[2]})
	}
	if len(color) != 6 {
		return 0, 0, 0, false
	}
	r, errR := strconv.ParseInt(color[0:2], 16, 32)
	g, errG := strconv.ParseInt(color[2:4], 16, 32)
	b, errB := strconv.ParseInt(color[4:6], 16, 32)
	if errR != nil || errG != nil || errB != nil {
		return 0, 0, 0, false
	}
	return int(r), int(g), int(b), true
}

// Output generates the PDF and outputs it according to dest parameter:
// "I" = inline HTTP response (web display)
// "D" = HTTP download with name
// "F" = save to file
// "S" = return as byte string
func (p *G3PDF) Output(dest, fileName string, isUTF8 bool) (interface{}, error) {
	dest = strings.ToUpper(strings.TrimSpace(dest))
	if dest == "" {
		dest = "S"
	}

	// Handle different output modes
	switch dest {
	case "I":
		// Return for inline display (will be handled by VM for HTTP response)
		buf := &bytes.Buffer{}
		if err := p.pdf.Output(buf); err != nil {
			return nil, fmt.Errorf("pdf generation failed: %w", err)
		}
		return buf.Bytes(), nil

	case "D":
		// Return for download
		if fileName == "" {
			fileName = "document.pdf"
		}
		buf := &bytes.Buffer{}
		if err := p.pdf.Output(buf); err != nil {
			return nil, fmt.Errorf("pdf generation failed: %w", err)
		}
		return map[string]interface{}{
			"content":  buf.Bytes(),
			"filename": fileName,
			"inline":   false,
		}, nil

	case "F":
		// Save to file
		if fileName == "" {
			fileName = "document.pdf"
		}
		return nil, p.pdf.OutputFileAndClose(fileName)

	case "S":
		// Return as string/bytes
		buf := &bytes.Buffer{}
		if err := p.pdf.Output(buf); err != nil {
			return nil, fmt.Errorf("pdf generation failed: %w", err)
		}
		return buf.Bytes(), nil

	default:
		return nil, fmt.Errorf("invalid output mode: %s, must be I, D, F, or S", dest)
	}
}

// WriteHTML renders an HTML string into the PDF
func (p *G3PDF) WriteHTML(html string) {
	if html == "" {
		return
	}

	if p.htmlState == nil {
		p.htmlState = &g3PDFHTMLState{
			styles: make([]*g3PDFHTMLStyle, 0, 10),
		}
	}

	// Reset state for new HTML rendering
	p.htmlState.tagStack = make([]string, 0, 20)
	p.htmlState.styles = make([]*g3PDFHTMLStyle, 0, 10)

	// Initialize rendering state
	p.htmlState.bold = false
	p.htmlState.italic = false
	p.htmlState.underline = false
	p.htmlState.boldCount = 0
	p.htmlState.italicCount = 0
	p.htmlState.underlineCount = 0
	p.htmlState.textColor = [3]int{0, 0, 0}
	p.htmlState.fontSize = 12
	p.htmlState.href = ""
	p.htmlState.inTable = false
	p.htmlState.inRow = false
	p.htmlState.rowStartY = 0
	p.htmlState.maxRowHeight = 0
	p.htmlState.tableColWidths = make(map[int]float64)
	p.htmlState.cellPadding = 1.2
	p.htmlState.cellSpacing = 0
	p.htmlState.colIndex = 0
	p.htmlState.cellText = ""
	p.htmlState.tdBegin = false
	p.htmlState.thBegin = false
	p.htmlState.tdWidth = 0
	p.htmlState.tdHeight = 0
	p.htmlState.tdAlign = "L"
	p.htmlState.tdBgColor = false
	p.htmlState.tdColorSet = false
	p.htmlState.currAlign = "L"
	p.htmlState.defaultFontSize = 12
	p.htmlState.colorSet = false
	p.htmlState.fontSet = false
	p.pdf.SetFont("helvetica", "", 12)
	p.pdf.SetTextColor(0, 0, 0)

	p.renderHTML(html)
}

// WriteHTMLFile loads an HTML file and renders it into the PDF
func (p *G3PDF) WriteHTMLFile(path string) error {
	content, err := os.ReadFile(path)
	if err != nil {
		return err
	}

	p.WriteHTML(string(content))
	return nil
}

// renderHTML is the main HTML parser and renderer
// It processes HTML tags and text content, maintaining state for font, color, and layout
func (p *G3PDF) renderHTML(input string) {
	input = strings.TrimSpace(input)
	pos := 0

	for pos < len(input) {
		// Find next tag
		tagStart := strings.Index(input[pos:], "<")
		if tagStart == -1 {
			// No more tags, output remaining text
			if pos < len(input) {
				p.handleHTMLText(input[pos:])
			}
			break
		}

		// Output text before tag
		if tagStart > 0 {
			p.handleHTMLText(input[pos : pos+tagStart])
		}

		// Find tag end
		tagEnd := strings.Index(input[pos+tagStart:], ">")
		if tagEnd == -1 {
			break
		}

		// Extract tag content
		tagContent := input[pos+tagStart : pos+tagStart+tagEnd+1]
		p.handleHTMLTag(tagContent)

		pos += tagStart + tagEnd + 1
	}
}

// handleHTMLText processes text content within HTML tags
func (p *G3PDF) handleHTMLText(text string) {
	if p.htmlState == nil {
		return
	}

	if text == "" {
		return
	}

	hadLeadingSpace := len(text) > 0 && (text[0] == ' ' || text[0] == '\t' || text[0] == '\n' || text[0] == '\r')
	hadTrailingSpace := len(text) > 0 && (text[len(text)-1] == ' ' || text[len(text)-1] == '\t' || text[len(text)-1] == '\n' || text[len(text)-1] == '\r')

	text = strings.ReplaceAll(text, "\r", "")
	text = strings.ReplaceAll(text, "\n", " ")
	for strings.Contains(text, "  ") {
		text = strings.ReplaceAll(text, "  ", " ")
	}
	trimmed := strings.TrimSpace(text)
	if trimmed == "" {
		if hadLeadingSpace || hadTrailingSpace {
			text = " "
		} else {
			return
		}
	} else {
		text = trimmed
		if hadLeadingSpace {
			text = " " + text
		}
		if hadTrailingSpace {
			text = text + " "
		}
	}

	if text == "" {
		return
	}

	if p.htmlState.inTable && (p.htmlState.tdBegin || p.htmlState.thBegin) {
		if p.htmlState.cellText != "" {
			p.htmlState.cellText += " "
		}
		p.htmlState.cellText += text
		return
	}

	p.updateHTMLStyle()

	if p.htmlState.href != "" {
		p.pdf.WriteLinkString(5, text, p.htmlState.href)
		return
	}

	p.pdf.Write(5, text)
}

// handleHTMLTag processes HTML tags
func (p *G3PDF) handleHTMLTag(tag string) {
	tag = strings.TrimSpace(tag)
	if !strings.HasPrefix(tag, "<") || !strings.HasSuffix(tag, ">") {
		return
	}

	content := tag[1 : len(tag)-1]
	isClosing := strings.HasPrefix(content, "/")

	if isClosing {
		tagName := strings.ToLower(strings.TrimSpace(content[1:]))
		p.handleHTMLCloseTag(tagName)
		p.popHTMLTagState(tagName)
	} else {
		tagName, attrs := parseHTMLTag(content)
		if tagName != "" {
			p.pushHTMLTagState(tagName)
			p.handleHTMLOpenTag(tagName, attrs)
		}
	}
}

// handleHTMLOpenTag handles opening tags
func (p *G3PDF) handleHTMLOpenTag(tagName string, attrs map[string]string) {
	if p.htmlState == nil {
		return
	}

	p.applyHTMLStyleAttributes(tagName, attrs)
	lineHeight := p.htmlLineHeight()
	left, _, _, _ := p.pdf.GetMargins()

	switch tagName {
	case "b", "strong":
		p.htmlState.bold = true

	case "i", "em":
		p.htmlState.italic = true

	case "u":
		p.htmlState.underline = true

	case "span", "font":
		return

	case "br":
		p.pdf.Ln(lineHeight)

	case "p":
		if p.pdf.GetX() > left+0.1 {
			p.pdf.Ln(lineHeight)
		}
		p.pdf.Ln(3)

	case "h1", "h2", "h3", "h4", "h5", "h6":
		if p.pdf.GetX() > left+0.1 {
			p.pdf.Ln(lineHeight)
		}
		size := 28 - (len(tagName) * 2)
		if size < 10 {
			size = 10
		}
		p.htmlState.fontSize = float64(size)
		p.htmlState.bold = true

	case "div", "section":
		if p.pdf.GetX() > left+0.1 {
			p.pdf.Ln(lineHeight)
		}

	case "hr":
		left, _, right, _ := p.pdf.GetMargins()
		pw, _ := p.pdf.GetPageSize()
		x1 := left
		x2 := pw - right
		p.pdf.Ln(2)
		y := p.pdf.GetY()
		p.pdf.Line(x1, y, x2, y)
		p.pdf.Ln(3)

	case "ul", "ol":
		p.pdf.Ln(2)
		p.htmlState.listDepth++
		p.htmlState.listStack = append(p.htmlState.listStack, &g3PDFHTMLListState{listType: tagName, listCount: 0})

	case "li":
		p.pdf.Ln(2)
		if len(p.htmlState.listStack) > 0 {
			current := p.htmlState.listStack[len(p.htmlState.listStack)-1]
			if current.listType == "ol" {
				current.listCount++
				p.pdf.Write(5, fmt.Sprintf("%d. ", current.listCount))
			} else {
				p.pdf.Write(5, "- ")
			}
		}

	case "table":
		p.pdf.Ln(3)
		p.htmlState.inTable = true
		p.htmlState.colIndex = 0
		p.htmlState.tableBorder = 1
		p.htmlState.tableColWidths = make(map[int]float64)

	case "tr":
		p.htmlState.inRow = true
		p.htmlState.colIndex = 0
		p.htmlState.rowStartY = p.pdf.GetY()
		p.htmlState.maxRowHeight = p.pdf.GetY()
		p.pdf.SetX(left)

	case "td", "th":
		p.htmlState.tdBegin = tagName == "td"
		p.htmlState.thBegin = tagName == "th"
		p.htmlState.cellText = ""
		p.htmlState.tdAlign = strings.ToUpper(strings.TrimSpace(attrs["align"]))
		if p.htmlState.tdAlign == "" {
			if tagName == "th" {
				p.htmlState.tdAlign = "C"
			} else {
				p.htmlState.tdAlign = "L"
			}
		}
		p.htmlState.tdWidth = parseHTMLNumericValue(attrs["width"])
		if p.htmlState.tdWidth > 0 {
			p.htmlState.tableColWidths[p.htmlState.colIndex] = p.htmlState.tdWidth
		}
		p.htmlState.tdHeight = parseHTMLNumericValue(attrs["height"])
		p.htmlState.tdBgColor = false
		p.htmlState.tdColorSet = false
		if tagName == "th" {
			p.htmlState.bold = true
			p.htmlState.tdBgColor = true
			p.htmlState.tdColorR = 230
			p.htmlState.tdColorG = 230
			p.htmlState.tdColorB = 230
		}
		if color := attrs["bgcolor"]; color != "" {
			if r, g, b, ok := parseHTMLColor(color); ok {
				p.htmlState.tdBgColor = true
				p.htmlState.tdColorR = float64(r)
				p.htmlState.tdColorG = float64(g)
				p.htmlState.tdColorB = float64(b)
			}
		}
		p.applyHTMLStyleAttributes(tagName, attrs)

	case "a":
		if href := attrs["href"]; href != "" {
			p.htmlState.href = href
		}

	case "img":
		src := attrs["src"]
		if src != "" {
			width := parseHTMLNumericValue(attrs["width"])
			height := parseHTMLNumericValue(attrs["height"])
			if width <= 0 {
				width = 40
			}
			if height <= 0 {
				height = 0
			}
			p.Image(src, math.NaN(), math.NaN(), width, height, "", "")
			p.pdf.Ln(8)
		}
	}
}

// handleHTMLCloseTag handles closing tags
func (p *G3PDF) handleHTMLCloseTag(tagName string) {
	if p.htmlState == nil {
		return
	}

	lineHeight := p.htmlLineHeight()
	left, _, _, _ := p.pdf.GetMargins()

	switch tagName {
	case "td", "th":
		p.renderHTMLTableCell()

	case "tr":
		if p.htmlState.maxRowHeight <= p.htmlState.rowStartY {
			fallbackHeight := lineHeight + (p.htmlState.cellPadding * 2)
			p.pdf.SetXY(left, p.htmlState.rowStartY+fallbackHeight)
		} else {
			p.pdf.SetXY(left, p.htmlState.maxRowHeight+p.htmlState.cellSpacing)
		}
		p.htmlState.inRow = false
		p.htmlState.colIndex = 0

	case "ul", "ol":
		if len(p.htmlState.listStack) > 0 {
			p.htmlState.listStack = p.htmlState.listStack[:len(p.htmlState.listStack)-1]
		}
		if p.htmlState.listDepth > 0 {
			p.htmlState.listDepth--
		}
		p.pdf.Ln(2)

	case "h1", "h2", "h3", "h4", "h5", "h6":
		p.pdf.Ln(6)

	case "p", "div", "section":
		p.pdf.Ln(5)

	case "table":
		p.htmlState.inTable = false
		p.pdf.Ln(5)
	}
}

// updateHTMLStyle applies current HTML styling to the PDF
func (p *G3PDF) updateHTMLStyle() {
	if p.htmlState == nil {
		return
	}

	style := ""
	if p.htmlState.bold {
		style += "B"
	}
	if p.htmlState.italic {
		style += "I"
	}
	if p.htmlState.underline {
		style += "U"
	}

	p.pdf.SetFont("helvetica", style, p.htmlState.fontSize)
	p.pdf.SetTextColor(p.htmlState.textColor[0], p.htmlState.textColor[1], p.htmlState.textColor[2])
}

// htmlLineHeight returns a safe text line height for the current HTML font size.
func (p *G3PDF) htmlLineHeight() float64 {
	if p.htmlState == nil || p.htmlState.fontSize <= 0 {
		return 5
	}
	lineHeight := p.htmlState.fontSize * 0.42
	if lineHeight < 5 {
		lineHeight = 5
	}
	return lineHeight
}

// pushHTMLTagState stores the current visual state so nested tags can restore cleanly.
func (p *G3PDF) pushHTMLTagState(tagName string) {
	if p.htmlState == nil {
		return
	}
	p.htmlState.tagStack = append(p.htmlState.tagStack, tagName)
	p.htmlState.styles = append(p.htmlState.styles, &g3PDFHTMLStyle{
		fontFamily: "helvetica",
		fontSize:   p.htmlState.fontSize,
		textColorR: float64(p.htmlState.textColor[0]),
		textColorG: float64(p.htmlState.textColor[1]),
		textColorB: float64(p.htmlState.textColor[2]),
		colorSet:   true,
		bold:       p.htmlState.bold,
		italic:     p.htmlState.italic,
		underline:  p.htmlState.underline,
		href:       p.htmlState.href,
	})
}

// popHTMLTagState restores the visual state that was active before an opening tag.
func (p *G3PDF) popHTMLTagState(tagName string) {
	if p.htmlState == nil || len(p.htmlState.styles) == 0 {
		return
	}
	last := p.htmlState.styles[len(p.htmlState.styles)-1]
	p.htmlState.styles = p.htmlState.styles[:len(p.htmlState.styles)-1]
	if len(p.htmlState.tagStack) > 0 {
		p.htmlState.tagStack = p.htmlState.tagStack[:len(p.htmlState.tagStack)-1]
	}
	p.htmlState.fontSize = last.fontSize
	p.htmlState.textColor = [3]int{int(last.textColorR), int(last.textColorG), int(last.textColorB)}
	p.htmlState.bold = last.bold
	p.htmlState.italic = last.italic
	p.htmlState.underline = last.underline
	p.htmlState.href = last.href
	p.updateHTMLStyle()
}

// applyHTMLStyleAttributes maps simple HTML attributes and inline CSS into current render state.
func (p *G3PDF) applyHTMLStyleAttributes(tagName string, attrs map[string]string) {
	if p.htmlState == nil {
		return
	}
	if colorValue := attrs["color"]; colorValue != "" {
		if r, g, b, ok := parseHTMLColor(colorValue); ok {
			p.htmlState.textColor = [3]int{r, g, b}
		}
	}
	if styleValue := attrs["style"]; styleValue != "" {
		styleMap := parseHTMLStyleDeclarations(styleValue)
		if colorValue := styleMap["color"]; colorValue != "" {
			if r, g, b, ok := parseHTMLColor(colorValue); ok {
				p.htmlState.textColor = [3]int{r, g, b}
			}
		}
		if tagName == "td" || tagName == "th" {
			if bgValue := styleMap["background-color"]; bgValue != "" {
				if r, g, b, ok := parseHTMLColor(bgValue); ok {
					p.htmlState.tdBgColor = true
					p.htmlState.tdColorR = float64(r)
					p.htmlState.tdColorG = float64(g)
					p.htmlState.tdColorB = float64(b)
				}
			}
		}
		if sizeValue := styleMap["font-size"]; sizeValue != "" {
			size := parseHTMLNumericValue(sizeValue)
			if size > 0 {
				p.htmlState.fontSize = size
			}
		}
	}
	if tagName == "font" {
		if sizeValue := attrs["size"]; sizeValue != "" {
			size := parseHTMLNumericValue(sizeValue)
			if size > 0 {
				p.htmlState.fontSize = size
			}
		}
	}
	p.updateHTMLStyle()
}

// renderHTMLTableCell outputs the current accumulated table cell and advances the cursor.
func (p *G3PDF) renderHTMLTableCell() {
	if p.htmlState == nil || !(p.htmlState.tdBegin || p.htmlState.thBegin) {
		return
	}
	left, _, right, _ := p.pdf.GetMargins()
	pageWidth, _ := p.pdf.GetPageSize()
	availableWidth := pageWidth - left - right

	width := p.htmlState.tdWidth
	if width <= 0 {
		if colWidth, ok := p.htmlState.tableColWidths[p.htmlState.colIndex]; ok && colWidth > 0 {
			width = colWidth
		} else {
			width = p.pdf.GetStringWidth(strings.TrimSpace(p.htmlState.cellText)) + (2 * p.htmlState.cellPadding)
			if width < 30 {
				width = 30
			}
			if width > availableWidth {
				width = availableWidth
			}
			p.htmlState.tableColWidths[p.htmlState.colIndex] = width
		}
	}

	height := p.htmlState.tdHeight
	if height <= 0 {
		height = 6
	}

	x := p.pdf.GetX()
	y := p.pdf.GetY()
	fill := p.htmlState.tdBgColor

	if fill {
		p.pdf.SetFillColor(int(p.htmlState.tdColorR), int(p.htmlState.tdColorG), int(p.htmlState.tdColorB))
	} else {
		p.pdf.SetFillColor(255, 255, 255)
	}

	if p.htmlState.tdColorSet {
		p.pdf.SetTextColor(int(p.htmlState.tdColorR), int(p.htmlState.tdColorG), int(p.htmlState.tdColorB))
	} else {
		p.pdf.SetTextColor(p.htmlState.textColor[0], p.htmlState.textColor[1], p.htmlState.textColor[2])
	}
	p.pdf.MultiCell(width, height, strings.TrimSpace(p.htmlState.cellText), "1", p.htmlState.tdAlign, fill)

	// Position cursor for next cell
	currY := p.pdf.GetY()
	p.pdf.SetXY(x+width, y)

	if currY > p.htmlState.maxRowHeight {
		p.htmlState.maxRowHeight = currY
	}

	p.htmlState.colIndex++
	p.htmlState.cellText = ""
	p.htmlState.tdBegin = false
	p.htmlState.thBegin = false
	p.htmlState.tdWidth = 0
	p.htmlState.tdHeight = 0
	p.htmlState.tdAlign = "L"
	p.htmlState.tdBgColor = false
	p.htmlState.tdColorSet = false
	p.updateHTMLStyle()
}

// parseHTMLTag splits an opening HTML tag into a normalized tag name and attributes.
func parseHTMLTag(content string) (string, map[string]string) {
	content = strings.TrimSpace(content)
	if content == "" {
		return "", nil
	}
	spaceIndex := strings.IndexAny(content, " \t\r\n")
	if spaceIndex == -1 {
		return strings.ToLower(content), map[string]string{}
	}
	tagName := strings.ToLower(strings.TrimSpace(content[:spaceIndex]))
	return tagName, parseHTMLAttributes(strings.TrimSpace(content[spaceIndex+1:]))
}

// parseHTMLAttributes parses quoted and unquoted attribute pairs from an HTML tag body.
func parseHTMLAttributes(input string) map[string]string {
	attrs := make(map[string]string)
	for len(input) > 0 {
		input = strings.TrimLeft(input, " \t\r\n")
		if input == "" {
			break
		}
		eqIndex := strings.IndexByte(input, '=')
		if eqIndex == -1 {
			attrs[strings.ToLower(strings.TrimSpace(input))] = ""
			break
		}
		key := strings.ToLower(strings.TrimSpace(input[:eqIndex]))
		input = strings.TrimLeft(input[eqIndex+1:], " \t\r\n")
		if input == "" {
			attrs[key] = ""
			break
		}
		quote := input[0]
		if quote == '\'' || quote == '"' {
			input = input[1:]
			endIndex := strings.IndexByte(input, quote)
			if endIndex == -1 {
				attrs[key] = input
				break
			}
			attrs[key] = input[:endIndex]
			input = input[endIndex+1:]
			continue
		}
		endIndex := strings.IndexAny(input, " \t\r\n")
		if endIndex == -1 {
			attrs[key] = input
			break
		}
		attrs[key] = input[:endIndex]
		input = input[endIndex+1:]
	}
	return attrs
}

// parseHTMLStyleDeclarations parses a CSS inline style string into a simple map.
func parseHTMLStyleDeclarations(input string) map[string]string {
	styleMap := make(map[string]string)
	parts := strings.Split(input, ";")
	for _, part := range parts {
		part = strings.TrimSpace(part)
		if part == "" {
			continue
		}
		pieces := strings.SplitN(part, ":", 2)
		if len(pieces) != 2 {
			continue
		}
		key := strings.ToLower(strings.TrimSpace(pieces[0]))
		value := strings.TrimSpace(pieces[1])
		styleMap[key] = value
	}
	return styleMap
}

// parseHTMLNumericValue extracts the numeric portion from a basic HTML size string.
func parseHTMLNumericValue(value string) float64 {
	value = strings.TrimSpace(value)
	if value == "" {
		return 0
	}
	for _, suffix := range []string{"px", "pt", "mm", "%"} {
		value = strings.TrimSuffix(strings.ToLower(value), suffix)
	}
	var parsed float64
	_, err := fmt.Sscanf(value, "%f", &parsed)
	if err != nil {
		return 0
	}
	return parsed
}

// parseHTMLColor parses a small set of HTML colors used by the PDF test suite.
func parseHTMLColor(value string) (int, int, int, bool) {
	return htmlColorToRGB(value)
}

// VM Integration Methods

// DispatchPropertyGet retrieves a property value for the VM
func (p *G3PDF) DispatchPropertyGet(name string) Value {
	propName := strings.ToLower(strings.TrimSpace(name))

	switch propName {
	case "lasterror":
		return NewString(p.lastError)
	case "page":
		return NewInteger(int64(p.pdf.PageNo()))
	case "x":
		return NewDouble(p.pdf.GetX())
	case "y":
		return NewDouble(p.pdf.GetY())
	case "w", "pagewidth":
		w, _ := p.pdf.GetPageSize()
		return NewDouble(w)
	case "h", "pageheight":
		_, h := p.pdf.GetPageSize()
		return NewDouble(h)
	case "version":
		return NewString("2.0 (go-pdf/fpdf)")
	default:
		return NewNull()
	}
}

// DispatchPropertySet sets a property value from the VM
func (p *G3PDF) DispatchPropertySet(name string, args []Value) bool {
	if len(args) == 0 {
		return false
	}

	propName := strings.ToLower(strings.TrimSpace(name))

	switch propName {
	case "x":
		p.SetX(toFloat(legacyValueToInterface(args[0], p.vm)))
		return true
	case "y":
		p.SetY(toFloat(legacyValueToInterface(args[0], p.vm)), true)
		return true
	}

	return false
}

// DispatchMethod invokes a method from the VM
func (p *G3PDF) DispatchMethod(name string, args []Value) Value {
	methodName := strings.ToLower(strings.TrimSpace(name))

	defer func() {
		if r := recover(); r != nil {
			p.setError(fmt.Sprintf("pdf error: %v", r))
		}
	}()

	// Document methods
	switch methodName {
	case "new", "init", "reset":
		orientation := "P"
		unit := "mm"
		size := "A4"

		if len(args) > 0 {
			orientation = fmt.Sprintf("%v", legacyValueToInterface(args[0], p.vm))
		}
		if len(args) > 1 {
			unit = fmt.Sprintf("%v", legacyValueToInterface(args[1], p.vm))
		}
		if len(args) > 2 {
			size = fmt.Sprintf("%v", legacyValueToInterface(args[2], p.vm))
		}

		p.Reset(orientation, unit, size)
		return legacyInterfaceToValue(p, p.vm)

	case "addpage":
		orientation := ""
		size := ""
		rotation := 0

		if len(args) > 0 {
			orientation = fmt.Sprintf("%v", legacyValueToInterface(args[0], p.vm))
		}
		if len(args) > 1 {
			size = fmt.Sprintf("%v", legacyValueToInterface(args[1], p.vm))
		}
		if len(args) > 2 {
			rotation = toInt(legacyValueToInterface(args[2], p.vm))
		}

		p.AddPage(orientation, size, rotation)
		return NewBool(true)

	case "close":
		p.Close()
		return NewBool(true)

	case "output":
		dest := "S"
		fileName := ""

		if len(args) > 0 {
			dest = strings.ToUpper(strings.TrimSpace(fmt.Sprintf("%v", legacyValueToInterface(args[0], p.vm))))
		}
		if len(args) > 1 {
			fileName = fmt.Sprintf("%v", legacyValueToInterface(args[1], p.vm))
		}

		// Generate PDF bytes
		buf := &bytes.Buffer{}
		if err := p.pdf.Output(buf); err != nil {
			p.setError(err.Error())
			return NewNull()
		}
		pdfBytes := buf.Bytes()

		switch dest {
		case "I", "D":
			// Write binary PDF directly to the HTTP response
			if p.vm != nil {
				resp := p.vm.host.Response()
				if dest == "D" {
					name := fileName
					if name == "" {
						name = "document.pdf"
					}
					resp.AddHeader("Content-Disposition", fmt.Sprintf("attachment; filename=\"%s\"", name))
				}
				resp.BinaryWrite(pdfBytes)
			}
			return NewBool(true)

		case "F":
			// Save to file path by writing generated bytes to disk
			if fileName == "" {
				fileName = "document.pdf"
			}
			if err := os.WriteFile(fileName, pdfBytes, 0666); err != nil {
				p.setError(err.Error())
				return NewBool(false)
			}
			return NewBool(true)

		default: // "S" – return bytes as binary string for Response.BinaryWrite in VBScript
			return NewString(string(pdfBytes))
		}

	// Font methods
	case "setfont":
		family := "helvetica"
		style := ""
		size := 12.0

		if len(args) > 0 {
			family = fmt.Sprintf("%v", legacyValueToInterface(args[0], p.vm))
		}
		if len(args) > 1 {
			style = fmt.Sprintf("%v", legacyValueToInterface(args[1], p.vm))
		}
		if len(args) > 2 {
			size = toFloat(legacyValueToInterface(args[2], p.vm))
		}

		p.SetFont(family, style, size)
		return NewBool(true)

	case "setfontsize":
		if len(args) > 0 {
			p.SetFontSize(toFloat(legacyValueToInterface(args[0], p.vm)))
		}
		return NewBool(true)

	// Color methods
	case "settextcolor":
		r, g, b := 0.0, math.NaN(), math.NaN()
		if len(args) > 0 {
			r = toFloat(legacyValueToInterface(args[0], p.vm))
		}
		if len(args) > 1 {
			g = toFloat(legacyValueToInterface(args[1], p.vm))
		}
		if len(args) > 2 {
			b = toFloat(legacyValueToInterface(args[2], p.vm))
		}
		p.SetTextColor(r, g, b)
		return NewBool(true)

	case "setdrawcolor":
		r, g, b := 0.0, math.NaN(), math.NaN()
		if len(args) > 0 {
			r = toFloat(legacyValueToInterface(args[0], p.vm))
		}
		if len(args) > 1 {
			g = toFloat(legacyValueToInterface(args[1], p.vm))
		}
		if len(args) > 2 {
			b = toFloat(legacyValueToInterface(args[2], p.vm))
		}
		p.SetDrawColor(r, g, b)
		return NewBool(true)

	case "setfillcolor":
		r, g, b := 0.0, math.NaN(), math.NaN()
		if len(args) > 0 {
			r = toFloat(legacyValueToInterface(args[0], p.vm))
		}
		if len(args) > 1 {
			g = toFloat(legacyValueToInterface(args[1], p.vm))
		}
		if len(args) > 2 {
			b = toFloat(legacyValueToInterface(args[2], p.vm))
		}
		p.SetFillColor(r, g, b)
		return NewBool(true)

	case "setlinewidth":
		if len(args) > 0 {
			p.SetLineWidth(toFloat(legacyValueToInterface(args[0], p.vm)))
		}
		return NewBool(true)

	// Margin methods
	case "setmargins":
		if len(args) < 2 {
			return NewBool(false)
		}
		left := toFloat(legacyValueToInterface(args[0], p.vm))
		top := toFloat(legacyValueToInterface(args[1], p.vm))
		var right *float64
		if len(args) > 2 {
			r := toFloat(legacyValueToInterface(args[2], p.vm))
			right = &r
		}
		p.SetMargins(left, top, right)
		return NewBool(true)

	case "setleftmargin":
		if len(args) > 0 {
			p.SetLeftMargin(toFloat(legacyValueToInterface(args[0], p.vm)))
		}
		return NewBool(true)

	case "settopmargin":
		if len(args) > 0 {
			p.SetTopMargin(toFloat(legacyValueToInterface(args[0], p.vm)))
		}
		return NewBool(true)

	case "setrightmargin":
		if len(args) > 0 {
			p.SetRightMargin(toFloat(legacyValueToInterface(args[0], p.vm)))
		}
		return NewBool(true)

	// Position methods
	case "setx":
		if len(args) > 0 {
			p.SetX(toFloat(legacyValueToInterface(args[0], p.vm)))
		}
		return NewBool(true)

	case "sety":
		resetX := true
		if len(args) > 0 {
			p.SetY(toFloat(legacyValueToInterface(args[0], p.vm)), resetX)
			if len(args) > 1 {
				resetX = toBool(legacyValueToInterface(args[1], p.vm))
				p.SetY(toFloat(legacyValueToInterface(args[0], p.vm)), resetX)
			}
		}
		return NewBool(true)

	case "setxy":
		if len(args) >= 2 {
			p.SetXY(
				toFloat(legacyValueToInterface(args[0], p.vm)),
				toFloat(legacyValueToInterface(args[1], p.vm)),
			)
		}
		return NewBool(true)

	case "getx":
		return NewDouble(p.pdf.GetX())

	case "gety":
		return NewDouble(p.pdf.GetY())

	case "ln":
		offset := -1.0
		if len(args) > 0 {
			offset = toFloat(legacyValueToInterface(args[0], p.vm))
		}
		p.Ln(offset)
		return NewBool(true)

	// Text methods
	case "cell":
		if len(args) < 1 {
			return NewBool(false)
		}
		w := toFloat(legacyValueToInterface(args[0], p.vm))
		h := 0.0
		txt := ""
		border := interface{}(0)
		ln := 0
		align := ""
		fill := false
		link := interface{}("")

		if len(args) > 1 {
			h = toFloat(legacyValueToInterface(args[1], p.vm))
		}
		if len(args) > 2 {
			txt = fmt.Sprintf("%v", legacyValueToInterface(args[2], p.vm))
		}
		if len(args) > 3 {
			border = legacyValueToInterface(args[3], p.vm)
		}
		if len(args) > 4 {
			ln = toInt(legacyValueToInterface(args[4], p.vm))
		}
		if len(args) > 5 {
			align = fmt.Sprintf("%v", legacyValueToInterface(args[5], p.vm))
		}
		if len(args) > 6 {
			fill = toBool(legacyValueToInterface(args[6], p.vm))
		}
		if len(args) > 7 {
			link = legacyValueToInterface(args[7], p.vm)
		}

		p.Cell(w, h, txt, border, ln, align, fill, link)
		return NewBool(true)

	case "multicell":
		if len(args) < 3 {
			return NewBool(false)
		}
		w := toFloat(legacyValueToInterface(args[0], p.vm))
		h := toFloat(legacyValueToInterface(args[1], p.vm))
		txt := fmt.Sprintf("%v", legacyValueToInterface(args[2], p.vm))
		border := interface{}(0)
		align := "J"
		fill := false

		if len(args) > 3 {
			border = legacyValueToInterface(args[3], p.vm)
		}
		if len(args) > 4 {
			align = fmt.Sprintf("%v", legacyValueToInterface(args[4], p.vm))
		}
		if len(args) > 5 {
			fill = toBool(legacyValueToInterface(args[5], p.vm))
		}

		p.MultiCell(w, h, txt, border, align, fill)
		return NewBool(true)

	case "write":
		if len(args) < 2 {
			return NewBool(false)
		}
		h := toFloat(legacyValueToInterface(args[0], p.vm))
		txt := fmt.Sprintf("%v", legacyValueToInterface(args[1], p.vm))
		link := interface{}("")

		if len(args) > 2 {
			link = legacyValueToInterface(args[2], p.vm)
		}

		p.Write(h, txt, link)
		return NewBool(true)

	case "text":
		if len(args) < 3 {
			return NewBool(false)
		}
		x := toFloat(legacyValueToInterface(args[0], p.vm))
		y := toFloat(legacyValueToInterface(args[1], p.vm))
		txt := fmt.Sprintf("%v", legacyValueToInterface(args[2], p.vm))

		p.Text(x, y, txt)
		return NewBool(true)

	// Drawing methods
	case "line":
		if len(args) < 4 {
			return NewBool(false)
		}
		x1 := toFloat(legacyValueToInterface(args[0], p.vm))
		y1 := toFloat(legacyValueToInterface(args[1], p.vm))
		x2 := toFloat(legacyValueToInterface(args[2], p.vm))
		y2 := toFloat(legacyValueToInterface(args[3], p.vm))

		p.Line(x1, y1, x2, y2)
		return NewBool(true)

	case "rect":
		if len(args) < 4 {
			return NewBool(false)
		}
		x := toFloat(legacyValueToInterface(args[0], p.vm))
		y := toFloat(legacyValueToInterface(args[1], p.vm))
		w := toFloat(legacyValueToInterface(args[2], p.vm))
		h := toFloat(legacyValueToInterface(args[3], p.vm))
		style := ""

		if len(args) > 4 {
			style = fmt.Sprintf("%v", legacyValueToInterface(args[4], p.vm))
		}

		p.Rect(x, y, w, h, style)
		return NewBool(true)

	// Image method
	case "image":
		if len(args) < 1 {
			return NewBool(false)
		}

		path := fmt.Sprintf("%v", legacyValueToInterface(args[0], p.vm))
		x, y := math.NaN(), math.NaN()
		w, h := 0.0, 0.0
		typ := ""
		link := interface{}("")

		if len(args) > 1 {
			x = toFloat(legacyValueToInterface(args[1], p.vm))
		}
		if len(args) > 2 {
			y = toFloat(legacyValueToInterface(args[2], p.vm))
		}
		if len(args) > 3 {
			w = toFloat(legacyValueToInterface(args[3], p.vm))
		}
		if len(args) > 4 {
			h = toFloat(legacyValueToInterface(args[4], p.vm))
		}
		if len(args) > 5 {
			typ = fmt.Sprintf("%v", legacyValueToInterface(args[5], p.vm))
		}
		if len(args) > 6 {
			link = legacyValueToInterface(args[6], p.vm)
		}

		p.Image(path, x, y, w, h, typ, link)
		return NewBool(true)

	// Link methods
	case "addlink":
		return NewInteger(int64(p.pdf.AddLink()))

	case "setlink":
		if len(args) < 1 {
			return NewBool(false)
		}
		linkID := toInt(legacyValueToInterface(args[0], p.vm))
		y := 0.0
		pageNum := 0

		if len(args) > 1 {
			y = toFloat(legacyValueToInterface(args[1], p.vm))
		}
		if len(args) > 2 {
			pageNum = toInt(legacyValueToInterface(args[2], p.vm))
		}

		p.pdf.SetLink(linkID, y, pageNum)
		return NewBool(true)

	case "link":
		if len(args) < 5 {
			return NewBool(false)
		}
		x := toFloat(legacyValueToInterface(args[0], p.vm))
		y := toFloat(legacyValueToInterface(args[1], p.vm))
		w := toFloat(legacyValueToInterface(args[2], p.vm))
		h := toFloat(legacyValueToInterface(args[3], p.vm))
		linkArg := legacyValueToInterface(args[4], p.vm)

		// Route to appropriate fpdf method based on link type
		switch lv := linkArg.(type) {
		case int:
			p.pdf.Link(x, y, w, h, lv)
		case int64:
			p.pdf.Link(x, y, w, h, int(lv))
		case float64:
			p.pdf.Link(x, y, w, h, int(lv))
		default:
			p.pdf.LinkString(x, y, w, h, fmt.Sprintf("%v", lv))
		}
		return NewBool(true)

	// Metadata methods
	case "settitle":
		isUTF8 := true
		if len(args) > 0 {
			p.title = fmt.Sprintf("%v", legacyValueToInterface(args[0], p.vm))
			p.pdf.SetTitle(p.title, isUTF8)
			if len(args) > 1 {
				isUTF8 = toBool(legacyValueToInterface(args[1], p.vm))
				p.pdf.SetTitle(p.title, isUTF8)
			}
		}
		return NewBool(true)

	case "setauthor":
		isUTF8 := true
		if len(args) > 0 {
			p.author = fmt.Sprintf("%v", legacyValueToInterface(args[0], p.vm))
			p.pdf.SetAuthor(p.author, isUTF8)
			if len(args) > 1 {
				isUTF8 = toBool(legacyValueToInterface(args[1], p.vm))
				p.pdf.SetAuthor(p.author, isUTF8)
			}
		}
		return NewBool(true)

	case "setsubject":
		isUTF8 := true
		if len(args) > 0 {
			p.subject = fmt.Sprintf("%v", legacyValueToInterface(args[0], p.vm))
			p.pdf.SetSubject(p.subject, isUTF8)
			if len(args) > 1 {
				isUTF8 = toBool(legacyValueToInterface(args[1], p.vm))
				p.pdf.SetSubject(p.subject, isUTF8)
			}
		}
		return NewBool(true)

	case "setkeywords":
		isUTF8 := true
		if len(args) > 0 {
			p.keywords = fmt.Sprintf("%v", legacyValueToInterface(args[0], p.vm))
			p.pdf.SetKeywords(p.keywords, isUTF8)
			if len(args) > 1 {
				isUTF8 = toBool(legacyValueToInterface(args[1], p.vm))
				p.pdf.SetKeywords(p.keywords, isUTF8)
			}
		}
		return NewBool(true)

	case "setcreator":
		isUTF8 := true
		if len(args) > 0 {
			p.creator = fmt.Sprintf("%v", legacyValueToInterface(args[0], p.vm))
			fullCreator := "AxonASP Librarie"
			if p.creator != "" {
				fullCreator = p.creator + " (AxonASP)"
			}
			p.pdf.SetCreator(fullCreator, isUTF8)
			if len(args) > 1 {
				isUTF8 = toBool(legacyValueToInterface(args[1], p.vm))
				p.pdf.SetCreator(fullCreator, isUTF8)
			}
		}
		return NewBool(true)

	case "aliasnbpages":
		if len(args) > 0 {
			p.aliasNbPages = fmt.Sprintf("%v", legacyValueToInterface(args[0], p.vm))
			p.pdf.AliasNbPages(p.aliasNbPages)
		}
		return NewBool(true)

	case "setdisplaymode":
		zoom := "default"
		layout := "default"
		if len(args) > 0 {
			zoom = fmt.Sprintf("%v", legacyValueToInterface(args[0], p.vm))
		}
		if len(args) > 1 {
			layout = fmt.Sprintf("%v", legacyValueToInterface(args[1], p.vm))
		}
		p.pdf.SetDisplayMode(zoom, layout)
		return NewBool(true)

	case "setcompression":
		if len(args) > 0 {
			p.compression = toBool(legacyValueToInterface(args[0], p.vm))
			p.pdf.SetCompression(p.compression)
		}
		return NewBool(true)

	// HTML methods
	case "writehtml", "html":
		if len(args) < 1 {
			return NewBool(false)
		}
		html := fmt.Sprintf("%v", legacyValueToInterface(args[0], p.vm))
		p.WriteHTML(html)
		return NewBool(true)

	case "writehtmlfile", "htmlfile", "loadhtmlfile":
		if len(args) < 1 {
			return NewBool(false)
		}
		path := fmt.Sprintf("%v", legacyValueToInterface(args[0], p.vm))
		err := p.WriteHTMLFile(path)
		if err != nil {
			p.setError(err.Error())
			return NewBool(false)
		}
		return NewBool(true)

	// Utility methods
	case "getpagewidth":
		w, _ := p.pdf.GetPageSize()
		return NewDouble(w)

	case "getpageheight":
		_, h := p.pdf.GetPageSize()
		return NewDouble(h)

	case "getstringwidth":
		if len(args) > 0 {
			txt := fmt.Sprintf("%v", legacyValueToInterface(args[0], p.vm))
			return NewDouble(p.pdf.GetStringWidth(txt))
		}
		return NewDouble(0)

	default:
		return NewNull()
	}
}
