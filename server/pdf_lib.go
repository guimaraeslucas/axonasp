/*
 * AxonASP Server
 * Copyright (C) 2026 G3pix Ltda. All rights reserved.
 *
 * Developed by Lucas Guimarães - G3pix Ltda
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
 * This is our translated version from PHP to Go of FPDF (v1.86),
 * originally authored by Olivier Plathey.
 */
package server

import (
	"bytes"
	"compress/zlib"
	"encoding/binary"
	"fmt"
	stdhtml "html"
	"image"
	_ "image/gif"
	stdjpeg "image/jpeg"
	_ "image/png"
	"io"
	"math"
	"os"
	"path/filepath"
	"regexp"
	"strconv"
	"strings"
	"time"
)

type pdfUVRange struct {
	start int
	count int
}

type pdfFont struct {
	typ       string
	name      string
	up        float64
	ut        float64
	cw        [256]int
	enc       string
	uv        map[int]interface{}
	subsetted bool
	n         int
	i         int
	file      string
	diff      string
}

type pdfImage struct {
	w    int
	h    int
	cs   string
	bpc  int
	f    string
	dp   string
	pal  []byte
	trns []int
	data []byte
	smk  []byte
	n    int
	i    int
}

type G3PDF struct {
	ctx *ExecutionContext

	state   int
	page    int
	n       int
	offsets map[int]int
	buffer  bytes.Buffer
	pages   map[int][]string

	compress bool
	k        float64

	defOrientation string
	curOrientation string
	stdPageSizes   map[string][2]float64
	defPageSize    [2]float64
	curPageSize    [2]float64
	curRotation    int
	pageInfo       map[int]map[string]interface{}

	wPt float64
	hPt float64
	w   float64
	h   float64

	lMargin float64
	tMargin float64
	rMargin float64
	bMargin float64
	cMargin float64

	x     float64
	y     float64
	lasth float64

	lineWidth float64
	fontpath  string

	coreFonts []string
	fonts     map[string]*pdfFont
	fontFiles map[string]map[string]int
	encodings map[string]int
	cmaps     map[string]int

	fontFamily  string
	fontStyle   string
	underline   bool
	currentFont *pdfFont
	fontSizePt  float64
	fontSize    float64

	drawColor string
	fillColor string
	textColor string
	colorFlag bool
	withAlpha bool
	ws        float64

	images map[string]*pdfImage

	pageLinks map[int][][]interface{}
	links     map[int][2]float64

	autoPageBreak    bool
	pageBreakTrigger float64
	inHeader         bool
	inFooter         bool
	aliasNbPages     string
	zoomMode         interface{}
	layoutMode       string
	metadata         map[string]string
	creationDate     time.Time
	pdfVersion       string

	assetFonts map[string]*pdfFont
	lastError  string
}

func NewG3PDF(ctx *ExecutionContext) *G3PDF {
	p := &G3PDF{ctx: ctx}
	p.Reset("P", "mm", "A4")
	return p
}

func (p *G3PDF) Reset(orientation, unit, size string) {
	p.state = 0
	p.page = 0
	p.n = 2
	p.offsets = map[int]int{}
	p.buffer.Reset()
	p.pages = map[int][]string{}
	p.pageInfo = map[int]map[string]interface{}{}
	p.fonts = map[string]*pdfFont{}
	p.fontFiles = map[string]map[string]int{}
	p.encodings = map[string]int{}
	p.cmaps = map[string]int{}
	p.images = map[string]*pdfImage{}
	p.links = map[int][2]float64{}
	p.pageLinks = map[int][][]interface{}{}
	p.inHeader = false
	p.inFooter = false
	p.lasth = 0
	p.fontFamily = ""
	p.fontStyle = ""
	p.fontSizePt = 12
	p.underline = false
	p.drawColor = "0 G"
	p.fillColor = "0 g"
	p.textColor = "0 g"
	p.colorFlag = false
	p.withAlpha = false
	p.ws = 0
	p.fontpath = ""
	p.coreFonts = []string{"courier", "helvetica", "times", "symbol", "zapfdingbats"}
	p.assetFonts = translatedFPDFFonts()

	switch strings.ToLower(strings.TrimSpace(unit)) {
	case "pt":
		p.k = 1
	case "mm":
		p.k = 72.0 / 25.4
	case "cm":
		p.k = 72.0 / 2.54
	case "in":
		p.k = 72
	default:
		p.setError("incorrect unit: " + unit)
		p.k = 72.0 / 25.4
	}

	p.stdPageSizes = map[string][2]float64{
		"a3":     {841.89 / p.k, 1190.55 / p.k},
		"a4":     {595.28 / p.k, 841.89 / p.k},
		"a5":     {420.94 / p.k, 595.28 / p.k},
		"letter": {612.0 / p.k, 792.0 / p.k},
		"legal":  {612.0 / p.k, 1008.0 / p.k},
	}

	sz := p.getPageSize(size)
	p.defPageSize = sz
	p.curPageSize = sz

	o := strings.ToLower(strings.TrimSpace(orientation))
	if o == "" || o == "p" || o == "portrait" {
		p.defOrientation = "P"
		p.w = sz[0]
		p.h = sz[1]
	} else {
		p.defOrientation = "L"
		p.w = sz[1]
		p.h = sz[0]
	}
	p.curOrientation = p.defOrientation
	p.wPt = p.w * p.k
	p.hPt = p.h * p.k
	p.curRotation = 0

	margin := 28.35 / p.k
	p.SetMargins(margin, margin, nil)
	p.cMargin = margin / 10
	p.lineWidth = 0.567 / p.k
	p.SetAutoPageBreak(true, 2*margin)
	p.SetDisplayMode("default", "default")
	p.SetCompression(true)
	p.metadata = map[string]string{"Producer": "G3Pix AxonASP Librarie"}
	p.pdfVersion = "1.3"
	p.creationDate = time.Now()
	p.lastError = ""
}

func (p *G3PDF) GetProperty(name string) interface{} {
	switch strings.ToLower(strings.TrimSpace(name)) {
	case "lasterror":
		return p.lastError
	case "fontpath":
		return p.fontpath
	case "page":
		return p.page
	case "x":
		return p.x
	case "y":
		return p.y
	case "w", "pagewidth":
		return p.w
	case "h", "pageheight":
		return p.h
	case "version":
		return "1.86"
	default:
		return nil
	}
}

func (p *G3PDF) SetProperty(name string, value interface{}) {
	switch strings.ToLower(strings.TrimSpace(name)) {
	case "fontpath":
		v := strings.TrimSpace(fmt.Sprintf("%v", value))
		if v != "" {
			p.fontpath = v
		}
	case "aliasnbpages":
		p.aliasNbPages = fmt.Sprintf("%v", value)
	case "x":
		p.SetX(toFloat(value))
	case "y":
		p.SetY(toFloat(value), true)
	}
}

func (p *G3PDF) CallMethod(name string, args ...interface{}) interface{} {
	m := strings.ToLower(strings.TrimSpace(name))
	defer func() {
		if r := recover(); r != nil {
			p.setError(fmt.Sprintf("pdf error: %v", r))
		}
	}()

	switch m {
	case "new", "init", "reset":
		o, u, s := "P", "mm", "A4"
		if len(args) > 0 {
			o = fmt.Sprintf("%v", args[0])
		}
		if len(args) > 1 {
			u = fmt.Sprintf("%v", args[1])
		}
		if len(args) > 2 {
			s = fmt.Sprintf("%v", args[2])
		}
		p.Reset(o, u, s)
		return p
	case "close":
		p.Close()
		return true
	case "addpage":
		o, s, r := "", "", 0
		if len(args) > 0 {
			o = fmt.Sprintf("%v", args[0])
		}
		if len(args) > 1 {
			s = fmt.Sprintf("%v", args[1])
		}
		if len(args) > 2 {
			r = toInt(args[2])
		}
		p.AddPage(o, s, r)
		return true
	case "setmargins":
		if len(args) < 2 {
			return nil
		}
		var right *float64
		if len(args) > 2 {
			r := toFloat(args[2])
			right = &r
		}
		p.SetMargins(toFloat(args[0]), toFloat(args[1]), right)
		return true
	case "setleftmargin":
		if len(args) > 0 {
			p.SetLeftMargin(toFloat(args[0]))
		}
		return true
	case "settopmargin":
		if len(args) > 0 {
			p.SetTopMargin(toFloat(args[0]))
		}
		return true
	case "setrightmargin":
		if len(args) > 0 {
			p.SetRightMargin(toFloat(args[0]))
		}
		return true
	case "setautopagebreak":
		if len(args) > 0 {
			margin := 0.0
			if len(args) > 1 {
				margin = toFloat(args[1])
			}
			p.SetAutoPageBreak(toBool(args[0]), margin)
		}
		return true
	case "setdisplaymode":
		zoom := interface{}("default")
		layout := "default"
		if len(args) > 0 {
			if s, ok := args[0].(string); ok {
				zoom = s
			} else {
				zoom = toFloat(args[0])
			}
		}
		if len(args) > 1 {
			layout = fmt.Sprintf("%v", args[1])
		}
		p.SetDisplayMode(zoom, layout)
		return true
	case "setcompression":
		if len(args) > 0 {
			p.SetCompression(toBool(args[0]))
		}
		return true
	case "settitle":
		if len(args) > 0 {
			utf := false
			if len(args) > 1 {
				utf = toBool(args[1])
			}
			p.SetTitle(fmt.Sprintf("%v", args[0]), utf)
		}
		return true
	case "setauthor":
		if len(args) > 0 {
			utf := false
			if len(args) > 1 {
				utf = toBool(args[1])
			}
			p.SetAuthor(fmt.Sprintf("%v", args[0]), utf)
		}
		return true
	case "setsubject":
		if len(args) > 0 {
			utf := false
			if len(args) > 1 {
				utf = toBool(args[1])
			}
			p.SetSubject(fmt.Sprintf("%v", args[0]), utf)
		}
		return true
	case "setkeywords":
		if len(args) > 0 {
			utf := false
			if len(args) > 1 {
				utf = toBool(args[1])
			}
			p.SetKeywords(fmt.Sprintf("%v", args[0]), utf)
		}
		return true
	case "setcreator":
		if len(args) > 0 {
			utf := false
			if len(args) > 1 {
				utf = toBool(args[1])
			}
			p.SetCreator(fmt.Sprintf("%v", args[0]), utf)
		}
		return true
	case "aliasnbpages":
		alias := "{nb}"
		if len(args) > 0 {
			alias = fmt.Sprintf("%v", args[0])
		}
		p.AliasNbPages(alias)
		return true
	case "setdrawcolor":
		if len(args) > 0 {
			g := math.NaN()
			b := math.NaN()
			if len(args) > 1 {
				g = toFloat(args[1])
			}
			if len(args) > 2 {
				b = toFloat(args[2])
			}
			p.SetDrawColor(toFloat(args[0]), g, b)
		}
		return true
	case "setfillcolor":
		if len(args) > 0 {
			g := math.NaN()
			b := math.NaN()
			if len(args) > 1 {
				g = toFloat(args[1])
			}
			if len(args) > 2 {
				b = toFloat(args[2])
			}
			p.SetFillColor(toFloat(args[0]), g, b)
		}
		return true
	case "settextcolor":
		if len(args) > 0 {
			g := math.NaN()
			b := math.NaN()
			if len(args) > 1 {
				g = toFloat(args[1])
			}
			if len(args) > 2 {
				b = toFloat(args[2])
			}
			p.SetTextColor(toFloat(args[0]), g, b)
		}
		return true
	case "getstringwidth":
		if len(args) == 0 {
			return float64(0)
		}
		return p.GetStringWidth(fmt.Sprintf("%v", args[0]))
	case "setlinewidth":
		if len(args) > 0 {
			p.SetLineWidth(toFloat(args[0]))
		}
		return true
	case "line":
		if len(args) >= 4 {
			p.Line(toFloat(args[0]), toFloat(args[1]), toFloat(args[2]), toFloat(args[3]))
		}
		return true
	case "rect":
		if len(args) >= 4 {
			style := ""
			if len(args) > 4 {
				style = fmt.Sprintf("%v", args[4])
			}
			p.Rect(toFloat(args[0]), toFloat(args[1]), toFloat(args[2]), toFloat(args[3]), style)
		}
		return true
	case "addfont":
		family, style, file, dir := "", "", "", ""
		if len(args) > 0 {
			family = fmt.Sprintf("%v", args[0])
		}
		if len(args) > 1 {
			style = fmt.Sprintf("%v", args[1])
		}
		if len(args) > 2 {
			file = fmt.Sprintf("%v", args[2])
		}
		if len(args) > 3 {
			dir = fmt.Sprintf("%v", args[3])
		}
		p.AddFont(family, style, file, dir)
		return true
	case "setfont":
		family, style, size := "", "", 0.0
		if len(args) > 0 {
			family = fmt.Sprintf("%v", args[0])
		}
		if len(args) > 1 {
			style = fmt.Sprintf("%v", args[1])
		}
		if len(args) > 2 {
			size = toFloat(args[2])
		}
		p.SetFont(family, style, size)
		return true
	case "setfontsize":
		if len(args) > 0 {
			p.SetFontSize(toFloat(args[0]))
		}
		return true
	case "addlink":
		return p.AddLink()
	case "setlink":
		if len(args) > 0 {
			y := 0.0
			page := -1
			if len(args) > 1 {
				y = toFloat(args[1])
			}
			if len(args) > 2 {
				page = toInt(args[2])
			}
			p.SetLink(toInt(args[0]), y, page)
		}
		return true
	case "link":
		if len(args) >= 5 {
			p.Link(toFloat(args[0]), toFloat(args[1]), toFloat(args[2]), toFloat(args[3]), args[4])
		}
		return true
	case "text":
		if len(args) >= 3 {
			p.Text(toFloat(args[0]), toFloat(args[1]), fmt.Sprintf("%v", args[2]))
		}
		return true
	case "cell":
		if len(args) > 0 {
			w := toFloat(args[0])
			h := 0.0
			txt := ""
			border := interface{}(0)
			ln := 0
			align := ""
			fill := false
			var link interface{} = ""
			if len(args) > 1 {
				h = toFloat(args[1])
			}
			if len(args) > 2 {
				txt = fmt.Sprintf("%v", args[2])
			}
			if len(args) > 3 {
				border = args[3]
			}
			if len(args) > 4 {
				ln = toInt(args[4])
			}
			if len(args) > 5 {
				align = fmt.Sprintf("%v", args[5])
			}
			if len(args) > 6 {
				fill = toBool(args[6])
			}
			if len(args) > 7 {
				link = args[7]
			}
			p.Cell(w, h, txt, border, ln, align, fill, link)
		}
		return true
	case "multicell":
		if len(args) >= 3 {
			border := interface{}(0)
			align := "J"
			fill := false
			if len(args) > 3 {
				border = args[3]
			}
			if len(args) > 4 {
				align = fmt.Sprintf("%v", args[4])
			}
			if len(args) > 5 {
				fill = toBool(args[5])
			}
			p.MultiCell(toFloat(args[0]), toFloat(args[1]), fmt.Sprintf("%v", args[2]), border, align, fill)
		}
		return true
	case "write":
		if len(args) >= 2 {
			link := interface{}("")
			if len(args) > 2 {
				link = args[2]
			}
			p.Write(toFloat(args[0]), fmt.Sprintf("%v", args[1]), link)
		}
		return true
	case "ln":
		if len(args) > 0 {
			p.Ln(toFloat(args[0]))
		} else {
			p.Ln(-1)
		}
		return true
	case "image":
		if len(args) > 0 {
			x, y, w, h := math.NaN(), math.NaN(), 0.0, 0.0
			typ := ""
			link := interface{}("")
			if len(args) > 1 {
				x = toFloat(args[1])
			}
			if len(args) > 2 {
				y = toFloat(args[2])
			}
			if len(args) > 3 {
				w = toFloat(args[3])
			}
			if len(args) > 4 {
				h = toFloat(args[4])
			}
			if len(args) > 5 {
				typ = fmt.Sprintf("%v", args[5])
			}
			if len(args) > 6 {
				link = args[6]
			}
			p.Image(fmt.Sprintf("%v", args[0]), x, y, w, h, typ, link)
		}
		return true
	case "getpagewidth":
		return p.w
	case "getpageheight":
		return p.h
	case "getx":
		return p.x
	case "setx":
		if len(args) > 0 {
			p.SetX(toFloat(args[0]))
		}
		return true
	case "gety":
		return p.y
	case "sety":
		if len(args) > 0 {
			resetX := true
			if len(args) > 1 {
				resetX = toBool(args[1])
			}
			p.SetY(toFloat(args[0]), resetX)
		}
		return true
	case "setxy":
		if len(args) >= 2 {
			p.SetXY(toFloat(args[0]), toFloat(args[1]))
		}
		return true
	case "output":
		dest, fileName := "", ""
		isUTF8 := false
		if len(args) > 0 {
			dest = fmt.Sprintf("%v", args[0])
		}
		if len(args) > 1 {
			fileName = fmt.Sprintf("%v", args[1])
		}
		if len(args) > 2 {
			isUTF8 = toBool(args[2])
		}
		res, err := p.Output(dest, fileName, isUTF8)
		if err != nil {
			p.setError(err.Error())
			return nil
		}
		return res
	case "writehtml", "html":
		if len(args) == 0 {
			return false
		}
		p.WriteHTML(fmt.Sprintf("%v", args[0]))
		return true
	case "writehtmlfile", "htmlfile", "loadhtmlfile":
		if len(args) == 0 {
			return false
		}
		err := p.WriteHTMLFile(fmt.Sprintf("%v", args[0]))
		if err != nil {
			p.setError(err.Error())
			return false
		}
		return true
	default:
		return nil
	}
}

type pdfHTMLStyle struct {
	colorR, colorG, colorB float64
	fontFamily             string
	fontStyle              string
	fontSize               float64
	colorSet               bool
}

type pdfHTMLState struct {
	p *G3PDF

	boldCount      int
	italicCount    int
	underlineCount int
	href           string
	pre            bool

	tableBorder int
	tdBegin     bool
	thBegin     bool
	tdWidth     float64
	tdHeight    float64
	tdAlign     string
	tdBgColor   bool
	trBgColor   bool
	cellPadding float64
	cellSpacing float64

	inTable        bool
	inRow          bool
	cellText       string
	colIndex       int
	tableColWidths map[int]float64
	rowStartY      float64
	maxRowHeight   float64
	tdWidthAttr    string

	// Cell-specific style overrides
	tdColorR, tdColorG, tdColorB float64
	tdColorSet                   bool

	styleStack []pdfHTMLStyle

	fontSet  bool
	colorSet bool

	listDepth int
	listType  string
	listCount int
	listStack []pdfHTMLListState
	currAlign string

	defaultFontSize float64
	scriptActive    bool
	scriptDeltaY    float64
}

func parseCSSStyle(style string) map[string]string {
	styles := map[string]string{}
	parts := strings.Split(style, ";")
	for _, part := range parts {
		kv := strings.SplitN(part, ":", 2)
		if len(kv) == 2 {
			key := strings.ToLower(strings.TrimSpace(kv[0]))
			val := strings.TrimSpace(kv[1])
			styles[key] = val
		}
	}
	return styles
}

type pdfHTMLListState struct {
	listType  string
	listCount int
}

func (p *G3PDF) WriteHTMLFile(path string) error {
	resolved := p.resolveHTMLPath(path)
	b, err := os.ReadFile(resolved)
	if err != nil {
		p.setError(err.Error())
		return err
	}
	p.WriteHTML(string(b))
	return nil
}

func (p *G3PDF) WriteHTML(htmlInput string) {
	if strings.TrimSpace(htmlInput) == "" {
		return
	}
	if p.page == 0 {
		p.AddPage("", "", 0)
	}

	state := &pdfHTMLState{
		p:               p,
		tdAlign:         "L",
		currAlign:       "L",
		defaultFontSize: p.fontSizePt,
		tableColWidths:  make(map[int]float64),
	}

	normalized := strings.ReplaceAll(htmlInput, "\r\n", "\n")
	normalized = strings.ReplaceAll(normalized, "\r", "\n")
	normalized = strings.ReplaceAll(normalized, "\t", "")

	state.renderHTML(normalized)
}

func (s *pdfHTMLState) pushStyle() {
	style := pdfHTMLStyle{
		fontFamily: s.p.fontFamily,
		fontStyle:  s.p.fontStyle,
		fontSize:   s.p.fontSizePt,
		colorSet:   s.colorSet,
	}
	// Note: G3PDF doesn't store raw RGB easily, but we can't easily retrieve them.
	// For now, we'll just track if a color was set.
	s.styleStack = append(s.styleStack, style)
}

func (s *pdfHTMLState) popStyle() {
	if len(s.styleStack) == 0 {
		return
	}
	last := s.styleStack[len(s.styleStack)-1]
	s.styleStack = s.styleStack[:len(s.styleStack)-1]

	s.p.SetFont(last.fontFamily, last.fontStyle, last.fontSize)
	if !last.colorSet {
		s.p.SetTextColor(0, math.NaN(), math.NaN()) // Reset to black
	}
	s.colorSet = last.colorSet
}

func (s *pdfHTMLState) renderHTML(input string) {
	tagRe := regexp.MustCompile(`(?is)<[^>]+>`)
	segments := tagRe.FindAllStringIndex(input, -1)
	pos := 0
	for _, seg := range segments {
		if seg[0] > pos {
			s.handleText(input[pos:seg[0]])
		}
		s.handleTag(input[seg[0]:seg[1]])
		pos = seg[1]
	}
	if pos < len(input) {
		s.handleText(input[pos:])
	}
}

func (s *pdfHTMLState) handleText(raw string) {
	if raw == "" {
		return
	}
	text := raw
	if !s.pre {
		// Replace all whitespace sequences with a single space
		re := regexp.MustCompile(`\s+`)
		text = re.ReplaceAllString(text, " ")
	}
	text = stdhtml.UnescapeString(text)
	text = normalizeHTMLTextForPDF(text)
	if text == "" {
		return
	}

	if s.href != "" {
		s.putLink(s.href, text)
		return
	}

	if s.tdBegin || s.thBegin {
		s.cellText += text
		return
	}

	if (s.inTable || s.inRow) && strings.TrimSpace(text) == "" {
		return
	}

	s.p.Write(5, text, "")
}

func (s *pdfHTMLState) handleTag(rawTag string) {
	tagContent := strings.TrimSpace(strings.TrimSuffix(strings.TrimPrefix(rawTag, "<"), ">"))
	if tagContent == "" {
		return
	}
	isClosing := strings.HasPrefix(tagContent, "/")
	isSelfClosing := strings.HasSuffix(tagContent, "/")

	if isClosing {
		tagName := strings.ToUpper(strings.TrimSpace(strings.TrimPrefix(tagContent, "/")))
		s.closeTag(tagName)
		return
	}

	if isSelfClosing {
		tagContent = strings.TrimSpace(strings.TrimSuffix(tagContent, "/"))
	}

	tagName, attrs := parseHTMLTag(tagContent)
	if tagName == "" {
		return
	}
	s.openTag(strings.ToUpper(tagName), attrs)

	if isSelfClosing {
		s.closeTag(strings.ToUpper(tagName))
	}
}

func (s *pdfHTMLState) openTag(tag string, attrs map[string]string) {
	// Handle CSS Style
	if style, ok := attrs["STYLE"]; ok {
		css := parseCSSStyle(style)
		if color, ok := css["color"]; ok {
			r, g, b := htmlColorToRGB(color)
			s.p.SetTextColor(float64(r), float64(g), float64(b))
			s.colorSet = true
		}
		if bgColor, ok := css["background-color"]; ok {
			r, g, b := htmlColorToRGB(bgColor)
			s.p.SetFillColor(float64(r), float64(g), float64(b))
			s.tdBgColor = true
		}
	}

	switch tag {
	case "STRONG", "B":
		s.setStyle("B", true)
	case "EM", "I":
		s.setStyle("I", true)
	case "U":
		s.setStyle("U", true)
	case "SUP":
		s.scriptActive = true
		s.scriptDeltaY = -2
		s.p.SetFont("", "", s.p.fontSizePt*0.7)
		s.p.SetXY(s.p.x, s.p.y+s.scriptDeltaY)
	case "SUB":
		s.scriptActive = true
		s.scriptDeltaY = 1
		s.p.SetFont("", "", s.p.fontSizePt*0.7)
		s.p.SetXY(s.p.x, s.p.y+s.scriptDeltaY)
	case "H1":
		s.headerTag(22, 150, 0, 0)
	case "H2":
		s.headerTag(18, 0, 0, 0)
	case "H3":
		s.headerTag(16, 0, 0, 0)
	case "H4":
		s.headerTag(14, 102, 0, 0)
	case "H5":
		s.headerTag(12, 0, 0, 150)
	case "H6":
		s.headerTag(10, 0, 0, 0)
	case "CENTER", "DIV", "P":
		s.p.Ln(5)
		if align, ok := attrs["ALIGN"]; ok {
			a := strings.ToUpper(align)
			switch a {
			case "CENTER":
				s.currAlign = "C"
			case "RIGHT":
				s.currAlign = "R"
			case "JUSTIFY":
				s.currAlign = "J"
			default:
				s.currAlign = "L"
			}
		}
	case "UL", "OL":
		s.p.Ln(2)
		s.listStack = append(s.listStack, pdfHTMLListState{listType: s.listType, listCount: s.listCount})
		if tag == "OL" {
			s.listType = "O"
		} else {
			s.listType = "U"
		}
		s.listCount = 0
		s.listDepth++
		s.p.SetLeftMargin(s.p.lMargin + 8)
		s.p.SetX(s.p.lMargin)
	case "LI":
		s.p.Ln(5)
		s.p.SetTextColor(100, 100, 100)
		marker := "\x95"
		if s.listType == "O" {
			s.listCount++
			marker = strconv.Itoa(s.listCount) + "."
		}
		s.p.Cell(6, 5, marker, 0, 0, "R", false, "")
		s.p.SetTextColor(0, math.NaN(), math.NaN())
	case "DL":
		s.p.Ln(3)
		s.p.SetLeftMargin(s.p.lMargin + 2)
		s.p.SetX(s.p.lMargin)
	case "DT":
		s.p.Ln(4)
		s.setStyle("B", true)
	case "DD":
		s.p.Ln(4)
		s.p.SetLeftMargin(s.p.lMargin + 8)
		s.p.SetX(s.p.lMargin)
	case "PRE":
		s.p.SetFont("courier", "", 10)
		s.pre = true
		s.p.Ln(4)
	case "BLOCKQUOTE":
		s.p.SetTextColor(80, 80, 80)
		s.p.Ln(3)
		s.p.SetLeftMargin(s.p.lMargin + 15)
		s.p.SetRightMargin(s.p.rMargin + 10)
		s.p.SetX(s.p.lMargin)
	case "HR":
		width := ""
		if v, ok := attrs["WIDTH"]; ok {
			width = v
		}
		if color, ok := attrs["COLOR"]; ok {
			r, g, b := htmlColorToRGB(color)
			s.p.SetDrawColor(float64(r), float64(g), float64(b))
		} else {
			s.p.SetDrawColor(200, 200, 200)
		}
		s.putLine(width)
		s.p.SetDrawColor(0, math.NaN(), math.NaN())
	case "BR":
		s.p.Ln(5)
	case "A":
		s.href = attrs["HREF"]
		s.p.SetTextColor(0, 0, 255)
		s.setStyle("U", true)
	case "IMG":
		src := attrs["SRC"]
		if src == "" {
			break
		}
		imgPath := s.p.resolveHTMLPath(src)
		w := parseHTMLLengthToMM(attrs["WIDTH"], s.p, 0)
		h := parseHTMLLengthToMM(attrs["HEIGHT"], s.p, 0)
		s.p.Image(imgPath, s.p.x, s.p.y, w, h, "", "")
		if !s.tdBegin && !s.thBegin {
			if h > 0 {
				s.p.Ln(h)
			} else {
				s.p.Ln(10)
			}
		}
	case "TABLE":
		s.p.Ln(5)
		s.inTable = true
		s.tableColWidths = make(map[int]float64)
		s.tableBorder = toInt(attrs["BORDER"])
		s.cellPadding = parseHTMLLengthToMM(attrs["CELLPADDING"], s.p, 1)
		s.cellSpacing = parseHTMLLengthToMM(attrs["CELLSPACING"], s.p, 0)
		if v, ok := attrs["BGCOLOR"]; ok {
			r, g, b := htmlColorToRGB(v)
			s.p.SetFillColor(float64(r), float64(g), float64(b))
		}
		if v, ok := attrs["COLOR"]; ok {
			r, g, b := htmlColorToRGB(v)
			s.p.SetTextColor(float64(r), float64(g), float64(b))
			s.colorSet = true
		}
	case "TR":
		// s.p.Ln(6) // REMOVED to avoid double spacing. closeTag TR handles the line jump.
		s.inRow = true
		s.colIndex = 0
		s.maxRowHeight = 0
		s.rowStartY = s.p.y
		s.trBgColor = false
		if v, ok := attrs["BGCOLOR"]; ok {
			r, g, b := htmlColorToRGB(v)
			s.p.SetFillColor(float64(r), float64(g), float64(b))
			s.trBgColor = true
		}
		if v, ok := attrs["COLOR"]; ok {
			r, g, b := htmlColorToRGB(v)
			s.p.SetTextColor(float64(r), float64(g), float64(b))
			s.colorSet = true
		}
	case "TD", "TH":
		s.colIndex++
		if tag == "TH" {
			s.setStyle("B", true)
			s.thBegin = true
		} else {
			s.tdBegin = true
		}
		s.cellText = ""
		s.tdWidthAttr = attrs["WIDTH"]
		s.tdHeight = parseHTMLLengthToMM(attrs["HEIGHT"], s.p, 6)

		// Apply cell spacing
		if s.cellSpacing > 0 {
			s.p.SetX(s.p.x + s.cellSpacing)
		}

		align := strings.ToUpper(attrs["ALIGN"])
		switch align {
		case "CENTER":
			s.tdAlign = "C"
		case "RIGHT":
			s.tdAlign = "R"
		case "JUSTIFY":
			s.tdAlign = "J"
		default:
			if s.thBegin {
				s.tdAlign = "C"
			} else {
				s.tdAlign = "L"
			}
		}

		s.tdBgColor = s.trBgColor
		if v, ok := attrs["BGCOLOR"]; ok {
			r, g, b := htmlColorToRGB(v)
			s.p.SetFillColor(float64(r), float64(g), float64(b))
			s.tdBgColor = true
		}

		if v, ok := attrs["COLOR"]; ok {
			r, g, b := htmlColorToRGB(v)
			s.p.SetTextColor(float64(r), float64(g), float64(b))
			s.tdColorR, s.tdColorG, s.tdColorB = float64(r), float64(g), float64(b)
			s.tdColorSet = true
			s.colorSet = true
		}

		// Apply padding by adjusting cMargin
		if s.cellPadding > 0 {
			s.p.cMargin = s.cellPadding
		}
	case "FONT":
		s.pushStyle()
		if color, ok := attrs["COLOR"]; ok {
			r, g, b := htmlColorToRGB(color)
			s.p.SetTextColor(float64(r), float64(g), float64(b))
			if s.tdBegin || s.thBegin {
				s.tdColorR, s.tdColorG, s.tdColorB = float64(r), float64(g), float64(b)
				s.tdColorSet = true
			}
			s.colorSet = true
		}
		if face, ok := attrs["FACE"]; ok {
			fontFace := strings.ToLower(face)
			if fontFace == "times" || fontFace == "courier" || fontFace == "helvetica" || fontFace == "symbol" {
				s.p.SetFont(fontFace, "", 0)
				s.fontSet = true
			}
		}
		if v, ok := attrs["SIZE"]; ok {
			size := toFloat(v)
			if size > 0 {
				s.p.SetFontSize(size)
			}
		}
	case "RED":
		s.p.SetTextColor(255, 0, 0)
	case "BLUE":
		s.p.SetTextColor(0, 0, 255)
	}
}

func (s *pdfHTMLState) closeTag(tag string) {
	if tag == "STRONG" {
		tag = "B"
	}
	if tag == "EM" {
		tag = "I"
	}

	switch tag {
	case "B", "I", "U":
		s.setStyle(tag, false)
	case "A":
		s.href = ""
		s.setStyle("U", false)
		s.p.SetTextColor(0, math.NaN(), math.NaN())
	case "SUP", "SUB":
		if s.scriptActive {
			s.p.SetFont("", "", s.defaultFontSize)
			s.p.SetXY(s.p.x, s.p.y-s.scriptDeltaY)
			s.scriptActive = false
		}
	case "H1", "H2", "H3", "H4", "H5", "H6":
		s.p.Ln(6)
		s.p.SetFont("helvetica", "", s.defaultFontSize)
		s.p.SetTextColor(0, math.NaN(), math.NaN())
		s.setStyle("B", false)
	case "P", "DIV", "CENTER":
		s.p.Ln(5)
		s.currAlign = "L"
	case "PRE":
		s.p.SetFont("helvetica", "", s.defaultFontSize)
		s.pre = false
	case "BLOCKQUOTE":
		s.p.SetLeftMargin(math.Max(0, s.p.lMargin-15))
		s.p.SetRightMargin(math.Max(0, s.p.rMargin-10))
		s.p.SetX(s.p.lMargin)
		s.p.Ln(5)
		s.p.SetTextColor(0, math.NaN(), math.NaN())
	case "UL", "OL":
		s.listDepth--
		s.p.SetLeftMargin(math.Max(0, s.p.lMargin-8))
		s.p.SetX(s.p.lMargin)
		s.p.Ln(2)
		if len(s.listStack) > 0 {
			last := s.listStack[len(s.listStack)-1]
			s.listStack = s.listStack[:len(s.listStack)-1]
			s.listType = last.listType
			s.listCount = last.listCount
		} else {
			s.listType = ""
			s.listCount = 0
		}
	case "DT":
		s.setStyle("B", false)
	case "DD":
		s.p.SetLeftMargin(math.Max(0, s.p.lMargin-8))
		s.p.SetX(s.p.lMargin)
	case "DL":
		s.p.SetLeftMargin(math.Max(0, s.p.lMargin-2))
		s.p.SetX(s.p.lMargin)
		s.p.Ln(2)
	case "TD", "TH":
		w := parseHTMLLengthToMM(s.tdWidthAttr, s.p, 0)
		if w == 0 {
			if cw, ok := s.tableColWidths[s.colIndex]; ok {
				w = cw
			} else {
				// Auto-grow for the first row
				w = s.p.GetStringWidth(s.cellText) + 2*s.p.cMargin
				if w < 30 {
					w = 30 // Minimum width
				}
				s.tableColWidths[s.colIndex] = w
			}
		} else {
			s.tableColWidths[s.colIndex] = w
		}

		h := s.tdHeight
		x, y := s.p.x, s.p.y

		// Apply cell-specific text color if set
		if s.tdColorSet {
			s.p.SetTextColor(s.tdColorR, s.tdColorG, s.tdColorB)
		}

		// Use MultiCell for the cell content
		s.p.MultiCell(w, h, s.cellText, s.tableBorder, s.tdAlign, s.tdBgColor)

		currY := s.p.y
		if currY > s.maxRowHeight {
			s.maxRowHeight = currY
		}

		// Restore position for next cell
		s.p.SetXY(x+w, y)

		if tag == "TH" {
			s.setStyle("B", false)
			s.thBegin = false
		} else {
			s.tdBegin = false
		}
		s.tdBgColor = false
		s.tdColorSet = false
		s.p.SetFillColor(255, math.NaN(), math.NaN())
		// Restore cMargin
		s.p.cMargin = (28.35 / s.p.k) / 10
	case "TR":
		if s.maxRowHeight > 0 {
			s.p.SetY(s.maxRowHeight+s.cellSpacing, true)
		} else {
			s.p.Ln(s.tdHeight + s.cellSpacing)
		}
		s.inRow = false
		s.trBgColor = false
	case "TABLE":
		s.tableBorder = 0
		s.inTable = false
		s.p.Ln(5)
		// Reset color to default after table
		s.p.SetTextColor(0, math.NaN(), math.NaN())
		s.colorSet = false
	case "FONT":
		s.popStyle()
		s.fontSet = false // Already handled by popStyle but keeping for safety
	case "RED", "BLUE":
		s.p.SetTextColor(0, math.NaN(), math.NaN())
		s.colorSet = false
	}
}

func (s *pdfHTMLState) headerTag(fontSize float64, r, g, b float64) {
	s.p.Ln(5)
	s.p.SetTextColor(r, g, b)
	s.p.SetFontSize(fontSize)
	s.setStyle("B", true)
}

func (s *pdfHTMLState) setStyle(tag string, enable bool) {
	switch tag {
	case "B":
		if enable {
			s.boldCount++
		} else if s.boldCount > 0 {
			s.boldCount--
		}
	case "I":
		if enable {
			s.italicCount++
		} else if s.italicCount > 0 {
			s.italicCount--
		}
	case "U":
		if enable {
			s.underlineCount++
		} else if s.underlineCount > 0 {
			s.underlineCount--
		}
	}

	style := ""
	if s.boldCount > 0 {
		style += "B"
	}
	if s.italicCount > 0 {
		style += "I"
	}
	if s.underlineCount > 0 {
		style += "U"
	}
	s.p.SetFont("", style, 0)
}

func (s *pdfHTMLState) putLink(url, text string) {
	s.p.SetTextColor(0, 0, 255)
	s.setStyle("U", true)
	s.p.Write(5, text, url)
	s.setStyle("U", false)
	s.p.SetTextColor(0, math.NaN(), math.NaN())
}

func (s *pdfHTMLState) putLine(width string) {
	lineW := parseHTMLLengthToMM(width, s.p, 0)
	if lineW <= 0 {
		lineW = s.p.w - s.p.lMargin - s.p.rMargin
	}
	s.p.Ln(2)
	x := s.p.x
	y := s.p.y
	s.p.Line(x, y, x+lineW, y)
	s.p.Ln(3)
}

func parseHTMLTag(content string) (string, map[string]string) {
	attrs := map[string]string{}
	content = strings.TrimSpace(content)
	if content == "" {
		return "", attrs
	}

	parts := strings.Fields(content)
	if len(parts) == 0 {
		return "", attrs
	}

	tagName := parts[0]
	rest := ""
	if len(content) > len(tagName) {
		rest = strings.TrimSpace(content[len(tagName):])
	}

	attrRe := regexp.MustCompile(`(?is)([a-zA-Z_:][-a-zA-Z0-9_:.]*)\s*=\s*("([^"]*)"|'([^']*)'|([^\s"'>]+))`)
	matches := attrRe.FindAllStringSubmatch(rest, -1)
	for _, m := range matches {
		key := strings.ToUpper(strings.TrimSpace(m[1]))
		val := ""
		switch {
		case m[3] != "":
			val = m[3]
		case m[4] != "":
			val = m[4]
		default:
			val = m[5]
		}
		attrs[key] = val
	}

	return tagName, attrs
}

func normalizeHTMLTextForPDF(text string) string {
	if text == "" {
		return text
	}

	text = strings.NewReplacer(
		"â€¢", "•",
		"â€“", "–",
		"â€”", "—",
		"â€˜", "‘",
		"â€™", "’",
		"â€œ", "“",
		"â€\x9d", "”",
		"â€¦", "…",
		"Â ", " ",
		"\u00A0", " ",
	).Replace(text)

	var b strings.Builder
	b.Grow(len(text))
	for _, r := range text {
		switch r {
		case '•':
			b.WriteByte(0x95)
		case '…':
			b.WriteByte(0x85)
		case '–':
			b.WriteByte(0x96)
		case '—':
			b.WriteByte(0x97)
		case '‘':
			b.WriteByte(0x91)
		case '’':
			b.WriteByte(0x92)
		case '“':
			b.WriteByte(0x93)
		case '”':
			b.WriteByte(0x94)
		case '™':
			b.WriteByte(0x99)
		case '€':
			b.WriteByte(0x80)
		case '\u00A0':
			b.WriteByte(' ')
		default:
			if r >= 0 && r <= 255 {
				b.WriteByte(byte(r))
			} else {
				b.WriteByte('?')
			}
		}
	}

	return b.String()
}

func htmlColorToRGB(color string) (int, int, int) {
	if color == "" {
		return 0, 0, 0
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
		return 0, 0, 0
	}
	r, errR := strconv.ParseInt(color[0:2], 16, 32)
	g, errG := strconv.ParseInt(color[2:4], 16, 32)
	b, errB := strconv.ParseInt(color[4:6], 16, 32)
	if errR != nil || errG != nil || errB != nil {
		return 0, 0, 0
	}
	return int(r), int(g), int(b)
}

func parseHTMLLengthToMM(v string, p *G3PDF, fallback float64) float64 {
	value := strings.TrimSpace(strings.ToLower(v))
	if value == "" {
		return fallback
	}
	if strings.HasSuffix(value, "%") {
		pct, err := strconv.ParseFloat(strings.TrimSuffix(value, "%"), 64)
		if err != nil {
			return fallback
		}
		return (p.w - p.lMargin - p.rMargin) * (pct / 100)
	}

	mult := 1.0
	switch {
	case strings.HasSuffix(value, "px"):
		value = strings.TrimSuffix(value, "px")
		mult = 25.4 / 72.0
	case strings.HasSuffix(value, "mm"):
		value = strings.TrimSuffix(value, "mm")
		mult = 1.0
	case strings.HasSuffix(value, "cm"):
		value = strings.TrimSuffix(value, "cm")
		mult = 10.0
	case strings.HasSuffix(value, "in"):
		value = strings.TrimSuffix(value, "in")
		mult = 25.4
	default:
		mult = 25.4 / 72.0
	}

	n, err := strconv.ParseFloat(strings.TrimSpace(value), 64)
	if err != nil {
		return fallback
	}
	return n * mult
}

func (p *G3PDF) resolveHTMLPath(path string) string {
	pTrim := strings.TrimSpace(path)
	if pTrim == "" {
		return pTrim
	}
	if filepath.IsAbs(pTrim) {
		return pTrim
	}
	if p.ctx != nil {
		mapped := p.ctx.Server_MapPath(pTrim)
		if mapped != "" {
			return mapped
		}
	}
	return pTrim
}

func (p *G3PDF) setError(msg string) {
	p.lastError = msg
}

func (p *G3PDF) panicError(msg string) {
	panic(fmt.Sprintf("G3PDF error: %s", msg))
}

func (p *G3PDF) SetMargins(left, top float64, right *float64) {
	p.lMargin = left
	p.tMargin = top
	if right == nil {
		p.rMargin = left
	} else {
		p.rMargin = *right
	}
}

func (p *G3PDF) SetLeftMargin(margin float64) {
	p.lMargin = margin
	if p.page > 0 && p.x < margin {
		p.x = margin
	}
}

func (p *G3PDF) SetTopMargin(margin float64)   { p.tMargin = margin }
func (p *G3PDF) SetRightMargin(margin float64) { p.rMargin = margin }

func (p *G3PDF) SetAutoPageBreak(auto bool, margin float64) {
	p.autoPageBreak = auto
	p.bMargin = margin
	p.pageBreakTrigger = p.h - margin
}

func (p *G3PDF) SetDisplayMode(zoom interface{}, layout string) {
	p.zoomMode = zoom
	p.layoutMode = strings.ToLower(layout)
}

func (p *G3PDF) SetCompression(compress bool)       { p.compress = compress }
func (p *G3PDF) SetTitle(title string, isUTF8 bool) { p.metadata["Title"] = p.metaText(title, isUTF8) }
func (p *G3PDF) SetAuthor(v string, isUTF8 bool)    { p.metadata["Author"] = p.metaText(v, isUTF8) }
func (p *G3PDF) SetSubject(v string, isUTF8 bool)   { p.metadata["Subject"] = p.metaText(v, isUTF8) }
func (p *G3PDF) SetKeywords(v string, isUTF8 bool)  { p.metadata["Keywords"] = p.metaText(v, isUTF8) }
func (p *G3PDF) SetCreator(v string, isUTF8 bool)   { p.metadata["Creator"] = p.metaText(v, isUTF8) }
func (p *G3PDF) AliasNbPages(alias string)          { p.aliasNbPages = alias }

func (p *G3PDF) metaText(v string, isUTF8 bool) string {
	if isUTF8 {
		return v
	}
	return latin1ToUTF8(v)
}

func (p *G3PDF) Close() {
	if p.state == 3 {
		return
	}
	if p.page == 0 {
		p.AddPage("", "", 0)
	}
	p.inFooter = true
	p.Footer()
	p.inFooter = false
	p.endPage()
	p.endDoc()
}

func (p *G3PDF) AddPage(orientation, size string, rotation int) {
	if p.state == 3 {
		p.panicError("the document is closed")
	}
	family := p.fontFamily
	style := p.fontStyle
	if p.underline {
		style += "U"
	}
	fontsize := p.fontSizePt
	lw := p.lineWidth
	dc := p.drawColor
	fc := p.fillColor
	tc := p.textColor
	cf := p.colorFlag
	if p.page > 0 {
		p.inFooter = true
		p.Footer()
		p.inFooter = false
		p.endPage()
	}
	p.beginPage(orientation, size, rotation)
	p.out("2 J")
	p.lineWidth = lw
	p.out(sprintf("%.2F w", lw*p.k))
	if family != "" {
		p.SetFont(family, style, fontsize)
	}
	p.drawColor = dc
	if dc != "0 G" {
		p.out(dc)
	}
	p.fillColor = fc
	if fc != "0 g" {
		p.out(fc)
	}
	p.textColor = tc
	p.colorFlag = cf

	p.inHeader = true
	p.Header()
	p.inHeader = false

	if p.lineWidth != lw {
		p.lineWidth = lw
		p.out(sprintf("%.2F w", lw*p.k))
	}
	if family != "" {
		p.SetFont(family, style, fontsize)
	}
	if p.drawColor != dc {
		p.drawColor = dc
		p.out(dc)
	}
	if p.fillColor != fc {
		p.fillColor = fc
		p.out(fc)
	}
	p.textColor = tc
	p.colorFlag = cf
}

func (p *G3PDF) Header()     {}
func (p *G3PDF) Footer()     {}
func (p *G3PDF) PageNo() int { return p.page }

func (p *G3PDF) SetDrawColor(r, g, b float64) {
	if math.IsNaN(g) || (r == 0 && g == 0 && b == 0) {
		p.drawColor = sprintf("%.3F G", r/255)
	} else {
		p.drawColor = sprintf("%.3F %.3F %.3F RG", r/255, g/255, b/255)
	}
	if p.page > 0 {
		p.out(p.drawColor)
	}
}

func (p *G3PDF) SetFillColor(r, g, b float64) {
	if math.IsNaN(g) || (r == 0 && g == 0 && b == 0) {
		p.fillColor = sprintf("%.3F g", r/255)
	} else {
		p.fillColor = sprintf("%.3F %.3F %.3F rg", r/255, g/255, b/255)
	}
	p.colorFlag = p.fillColor != p.textColor
	if p.page > 0 {
		p.out(p.fillColor)
	}
}

func (p *G3PDF) SetTextColor(r, g, b float64) {
	if math.IsNaN(g) || (r == 0 && g == 0 && b == 0) {
		p.textColor = sprintf("%.3F g", r/255)
	} else {
		p.textColor = sprintf("%.3F %.3F %.3F rg", r/255, g/255, b/255)
	}
	p.colorFlag = p.fillColor != p.textColor
}

func (p *G3PDF) GetStringWidth(s string) float64 {
	if p.currentFont == nil {
		return 0
	}
	w := 0
	for _, c := range []byte(s) {
		w += p.currentFont.cw[c]
	}
	return float64(w) * p.fontSize / 1000
}

func (p *G3PDF) SetLineWidth(width float64) {
	p.lineWidth = width
	if p.page > 0 {
		p.out(sprintf("%.2F w", width*p.k))
	}
}

func (p *G3PDF) Line(x1, y1, x2, y2 float64) {
	p.out(sprintf("%.2F %.2F m %.2F %.2F l S", x1*p.k, (p.h-y1)*p.k, x2*p.k, (p.h-y2)*p.k))
}

func (p *G3PDF) Rect(x, y, w, h float64, style string) {
	op := "S"
	switch style {
	case "F":
		op = "f"
	case "FD", "DF":
		op = "B"
	}
	p.out(sprintf("%.2F %.2F %.2F %.2F re %s", x*p.k, (p.h-y)*p.k, w*p.k, -h*p.k, op))
}

func (p *G3PDF) AddFont(family, style, file, dir string) {
	family = strings.ToLower(strings.TrimSpace(family))
	if file == "" {
		file = strings.ReplaceAll(family, " ", "") + strings.ToLower(style) + ".php"
	}
	style = strings.ToUpper(style)
	if style == "IB" {
		style = "BI"
	}
	fontkey := family + style
	if _, ok := p.fonts[fontkey]; ok {
		return
	}
	if strings.Contains(file, "/") || strings.Contains(file, "\\") {
		p.panicError("incorrect font definition file name: " + file)
	}
	if dir == "" {
		dir = p.fontpath
	}
	info, ok := p.loadFontAsset(file)
	if !ok {
		p.panicError("could not load embedded font definition: " + file)
	}
	clone := *info
	clone.i = len(p.fonts) + 1
	p.fonts[fontkey] = &clone
}

func (p *G3PDF) SetFont(family, style string, size float64) {
	if family == "" {
		family = p.fontFamily
	} else {
		family = strings.ToLower(family)
	}
	style = strings.ToUpper(style)
	if strings.Contains(style, "U") {
		p.underline = true
		style = strings.ReplaceAll(style, "U", "")
	} else {
		p.underline = false
	}
	if style == "IB" {
		style = "BI"
	}
	if size == 0 {
		size = p.fontSizePt
	}
	if p.fontFamily == family && p.fontStyle == style && p.fontSizePt == size {
		return
	}
	fontkey := family + style
	if _, ok := p.fonts[fontkey]; !ok {
		if family == "arial" {
			family = "helvetica"
		}
		if containsString(p.coreFonts, family) {
			if family == "symbol" || family == "zapfdingbats" {
				style = ""
			}
			fontkey = family + style
			if _, ok2 := p.fonts[fontkey]; !ok2 {
				p.AddFont(family, style, "", "")
			}
		} else {
			p.panicError("undefined font: " + family + " " + style)
		}
	}
	p.fontFamily = family
	p.fontStyle = style
	p.fontSizePt = size
	p.fontSize = size / p.k
	p.currentFont = p.fonts[fontkey]
	if p.page > 0 {
		p.out(sprintf("BT /F%d %.2F Tf ET", p.currentFont.i, p.fontSizePt))
	}
}

func (p *G3PDF) SetFontSize(size float64) {
	if p.fontSizePt == size {
		return
	}
	p.fontSizePt = size
	p.fontSize = size / p.k
	if p.page > 0 && p.currentFont != nil {
		p.out(sprintf("BT /F%d %.2F Tf ET", p.currentFont.i, p.fontSizePt))
	}
}

func (p *G3PDF) AddLink() int {
	n := len(p.links) + 1
	p.links[n] = [2]float64{0, 0}
	return n
}

func (p *G3PDF) SetLink(link int, y float64, page int) {
	if y == -1 {
		y = p.y
	}
	if page == -1 {
		page = p.page
	}
	p.links[link] = [2]float64{float64(page), y}
}

func (p *G3PDF) Link(x, y, w, h float64, link interface{}) {
	p.pageLinks[p.page] = append(p.pageLinks[p.page], []interface{}{x * p.k, p.hPt - y*p.k, w * p.k, h * p.k, link})
}

func (p *G3PDF) Text(x, y float64, txt string) {
	if p.currentFont == nil {
		p.panicError("no font has been set")
	}
	s := sprintf("BT %.2F %.2F Td (%s) Tj ET", x*p.k, (p.h-y)*p.k, p.escape(txt))
	if p.underline && txt != "" {
		s += " " + p.doUnderline(x, y, txt)
	}
	if p.colorFlag {
		s = "q " + p.textColor + " " + s + " Q"
	}
	p.out(s)
}

func (p *G3PDF) AcceptPageBreak() bool { return p.autoPageBreak }

func (p *G3PDF) Cell(w, h float64, txt string, border interface{}, ln int, align string, fill bool, link interface{}) {
	k := p.k
	if p.y+h > p.pageBreakTrigger && !p.inHeader && !p.inFooter && p.AcceptPageBreak() {
		x := p.x
		ws := p.ws
		if ws > 0 {
			p.ws = 0
			p.out("0 Tw")
		}
		p.AddPage(p.curOrientation, "", p.curRotation)
		p.x = x
		if ws > 0 {
			p.ws = ws
			p.out(sprintf("%.3F Tw", ws*k))
		}
	}
	if w == 0 {
		w = p.w - p.rMargin - p.x
	}
	s := ""
	if fill || border == 1 || border == "1" {
		op := "S"
		if fill {
			if border == 1 || border == "1" {
				op = "B"
			} else {
				op = "f"
			}
		}
		s = sprintf("%.2F %.2F %.2F %.2F re %s ", p.x*k, (p.h-p.y)*k, w*k, -h*k, op)
	}
	if bs, ok := border.(string); ok {
		x := p.x
		y := p.y
		if strings.Contains(bs, "L") {
			s += sprintf("%.2F %.2F m %.2F %.2F l S ", x*k, (p.h-y)*k, x*k, (p.h-(y+h))*k)
		}
		if strings.Contains(bs, "T") {
			s += sprintf("%.2F %.2F m %.2F %.2F l S ", x*k, (p.h-y)*k, (x+w)*k, (p.h-y)*k)
		}
		if strings.Contains(bs, "R") {
			s += sprintf("%.2F %.2F m %.2F %.2F l S ", (x+w)*k, (p.h-y)*k, (x+w)*k, (p.h-(y+h))*k)
		}
		if strings.Contains(bs, "B") {
			s += sprintf("%.2F %.2F m %.2F %.2F l S ", x*k, (p.h-(y+h))*k, (x+w)*k, (p.h-(y+h))*k)
		}
	}
	if txt != "" {
		if p.currentFont == nil {
			p.panicError("no font has been set")
		}
		dx := p.cMargin
		switch align {
		case "R":
			dx = w - p.cMargin - p.GetStringWidth(txt)
		case "C":
			dx = (w - p.GetStringWidth(txt)) / 2
		}
		if p.colorFlag {
			s += "q " + p.textColor + " "
		}
		s += sprintf("BT %.2F %.2F Td (%s) Tj ET", (p.x+dx)*k, (p.h-(p.y+0.5*h+0.3*p.fontSize))*k, p.escape(txt))
		if p.underline {
			s += " " + p.doUnderline(p.x+dx, p.y+0.5*h+0.3*p.fontSize, txt)
		}
		if p.colorFlag {
			s += " Q"
		}
		if link != "" && link != nil {
			p.Link(p.x+dx, p.y+0.5*h-0.5*p.fontSize, p.GetStringWidth(txt), p.fontSize, link)
		}
	}
	if s != "" {
		p.out(s)
	}
	p.lasth = h
	if ln > 0 {
		p.y += h
		if ln == 1 {
			p.x = p.lMargin
		}
	} else {
		p.x += w
	}
}

func (p *G3PDF) MultiCell(w, h float64, txt string, border interface{}, align string, fill bool) {
	if p.currentFont == nil {
		p.panicError("no font has been set")
	}
	if w == 0 {
		w = p.w - p.rMargin - p.x
	}
	wmax := (w - 2*p.cMargin) * 1000 / p.fontSize
	s := strings.ReplaceAll(txt, "\r", "")
	nb := len(s)
	if nb > 0 && s[nb-1] == '\n' {
		nb--
	}
	b := ""
	b2 := ""
	if border != nil && border != 0 && border != "0" && border != "" {
		if border == 1 || border == "1" {
			b = "LRT"
			b2 = "LR"
		} else if bs, ok := border.(string); ok {
			if strings.Contains(bs, "L") {
				b2 += "L"
			}
			if strings.Contains(bs, "R") {
				b2 += "R"
			}
			if strings.Contains(bs, "T") {
				b = b2 + "T"
			} else {
				b = b2
			}
		}
	}
	sep := -1
	i, j := 0, 0
	l, ns, nl := 0, 0, 1
	for i < nb {
		c := s[i]
		if c == '\n' {
			if p.ws > 0 {
				p.ws = 0
				p.out("0 Tw")
			}
			p.Cell(w, h, s[j:i], b, 2, align, fill, "")
			i++
			sep = -1
			j = i
			l = 0
			ns = 0
			nl++
			if b != "" && nl == 2 {
				b = b2
			}
			continue
		}
		if c == ' ' {
			sep = i
			ns++
		}
		l += p.charWidth(c)
		if float64(l) > wmax {
			if sep == -1 {
				if i == j {
					i++
				}
				if p.ws > 0 {
					p.ws = 0
					p.out("0 Tw")
				}
				p.Cell(w, h, s[j:i], b, 2, align, fill, "")
			} else {
				if align == "J" {
					spaces := strings.Count(s[j:sep], " ")
					if spaces > 0 {
						strW := p.GetStringWidth(s[j:sep])
						p.ws = (w - 2*p.cMargin - strW) / float64(spaces)
						p.out(sprintf("%.3F Tw", p.ws*p.k))
					}
				}
				p.Cell(w, h, s[j:sep], b, 2, align, fill, "")
				i = sep + 1
			}
			sep = -1
			j = i
			l = 0
			ns = 0
			nl++
			if b != "" && nl == 2 {
				b = b2
			}
		} else {
			i++
		}
	}
	if p.ws > 0 {
		p.ws = 0
		p.out("0 Tw")
	}
	if border == 1 || border == "1" {
		b += "B"
	} else if bs, ok := border.(string); ok && strings.Contains(bs, "B") {
		b += "B"
	}
	p.Cell(w, h, s[j:i], b, 2, align, fill, "")
	p.x = p.lMargin
}

func (p *G3PDF) Write(h float64, txt string, link interface{}) {
	if p.currentFont == nil {
		p.panicError("no font has been set")
	}
	w := p.w - p.rMargin - p.x
	wmax := (w - 2*p.cMargin) * 1000 / p.fontSize
	s := strings.ReplaceAll(txt, "\r", "")
	nb := len(s)
	sep := -1
	i, j, l, nl := 0, 0, 0, 1
	for i < nb {
		c := s[i]
		if c == '\n' {
			p.Cell(w, h, s[j:i], 0, 2, "", false, link)
			i++
			sep = -1
			j = i
			l = 0
			if nl == 1 {
				p.x = p.lMargin
				w = p.w - p.rMargin - p.x
				wmax = (w - 2*p.cMargin) * 1000 / p.fontSize
			}
			nl++
			continue
		}
		if c == ' ' {
			sep = i
		}
		l += p.charWidth(c)
		if float64(l) > wmax {
			if sep == -1 {
				if p.x > p.lMargin {
					p.x = p.lMargin
					p.y += h
					w = p.w - p.rMargin - p.x
					wmax = (w - 2*p.cMargin) * 1000 / p.fontSize
					i++
					nl++
					continue
				}
				if i == j {
					i++
				}
				p.Cell(w, h, s[j:i], 0, 2, "", false, link)
			} else {
				p.Cell(w, h, s[j:sep], 0, 2, "", false, link)
				i = sep + 1
			}
			sep = -1
			j = i
			l = 0
			if nl == 1 {
				p.x = p.lMargin
				w = p.w - p.rMargin - p.x
				wmax = (w - 2*p.cMargin) * 1000 / p.fontSize
			}
			nl++
		} else {
			i++
		}
	}
	if i != j {
		p.Cell(float64(l)/1000*p.fontSize, h, s[j:], 0, 0, "", false, link)
	}
}

func (p *G3PDF) Ln(h float64) {
	p.x = p.lMargin
	if h < 0 {
		p.y += p.lasth
	} else {
		p.y += h
	}
}

func (p *G3PDF) Image(file string, x, y, w, h float64, typ string, link interface{}) {
	if file == "" {
		p.panicError("image file name is empty")
	}
	info, ok := p.images[file]
	if !ok {
		if typ == "" {
			ext := strings.ToLower(strings.TrimPrefix(filepath.Ext(file), "."))
			if ext == "" {
				p.panicError("image file has no extension and no type was specified: " + file)
			}
			typ = ext
		}
		typ = strings.ToLower(typ)
		if typ == "jpeg" {
			typ = "jpg"
		}
		switch typ {
		case "jpg":
			info = p.parseImageFile(file)
		case "png", "gif":
			info = p.parseImageFile(file)
		default:
			p.panicError("unsupported image type: " + typ)
		}
		info.i = len(p.images) + 1
		p.images[file] = info
	}

	if w == 0 && h == 0 {
		w = -96
		h = -96
	}
	if w < 0 {
		w = -float64(info.w) * 72 / w / p.k
	}
	if h < 0 {
		h = -float64(info.h) * 72 / h / p.k
	}
	if w == 0 {
		w = h * float64(info.w) / float64(info.h)
	}
	if h == 0 {
		h = w * float64(info.h) / float64(info.w)
	}
	if math.IsNaN(y) {
		if p.y+h > p.pageBreakTrigger && !p.inHeader && !p.inFooter && p.AcceptPageBreak() {
			x2 := p.x
			p.AddPage(p.curOrientation, "", p.curRotation)
			p.x = x2
		}
		y = p.y
		p.y += h
	}
	if math.IsNaN(x) {
		x = p.x
	}
	p.out(sprintf("q %.2F 0 0 %.2F %.2F %.2F cm /I%d Do Q", w*p.k, h*p.k, x*p.k, (p.h-(y+h))*p.k, info.i))
	if link != "" && link != nil {
		p.Link(x, y, w, h, link)
	}
}

func (p *G3PDF) SetX(x float64) {
	if x >= 0 {
		p.x = x
	} else {
		p.x = p.w + x
	}
}

func (p *G3PDF) SetY(y float64, resetX bool) {
	if y >= 0 {
		p.y = y
	} else {
		p.y = p.h + y
	}
	if resetX {
		p.x = p.lMargin
	}
}

func (p *G3PDF) SetXY(x, y float64) {
	p.SetX(x)
	p.SetY(y, false)
}

func (p *G3PDF) Output(dest, name string, isUTF8 bool) (string, error) {
	p.Close()
	if len(name) == 1 && len(dest) != 1 {
		tmp := dest
		dest = name
		name = tmp
	}
	if dest == "" {
		dest = "I"
	}
	if name == "" {
		name = "doc.pdf"
	}
	pdf := p.buffer.Bytes()
	switch strings.ToUpper(dest) {
	case "I", "D":
		if p.ctx != nil && p.ctx.Response != nil {
			p.ctx.Response.SetContentType("application/pdf")
			disp := "inline"
			if strings.ToUpper(dest) == "D" {
				disp = "attachment"
			}
			p.ctx.Response.AddHeader("Content-Disposition", disp+"; "+p.httpEncode("filename", name, isUTF8))
			p.ctx.Response.AddHeader("Cache-Control", "private, max-age=0, must-revalidate")
			if err := p.ctx.Response.BinaryWrite(pdf); err != nil {
				return "", err
			}
			return "", nil
		}
		return string(pdf), nil
	case "F":
		if err := os.WriteFile(name, pdf, 0644); err != nil {
			return "", err
		}
		return "", nil
	case "S":
		return string(pdf), nil
	default:
		return "", fmt.Errorf("incorrect output destination: %s", dest)
	}
}

func (p *G3PDF) getPageSize(size string) [2]float64 {
	s := strings.ToLower(strings.TrimSpace(size))
	if s == "" {
		s = "a4"
	}
	if v, ok := p.stdPageSizes[s]; ok {
		return v
	}
	return p.stdPageSizes["a4"]
}

func (p *G3PDF) beginPage(orientation, size string, rotation int) {
	p.page++
	p.pages[p.page] = []string{}
	p.pageLinks[p.page] = [][]any{}
	p.state = 2
	p.x = p.lMargin
	p.y = p.tMargin
	p.fontFamily = ""

	if orientation == "" {
		orientation = p.defOrientation
	} else {
		orientation = strings.ToUpper(string(orientation[0]))
	}

	var ps [2]float64
	if size == "" {
		ps = p.defPageSize
	} else {
		ps = p.getPageSize(size)
	}
	if orientation != p.curOrientation || ps != p.curPageSize {
		if orientation == "P" {
			p.w, p.h = ps[0], ps[1]
		} else {
			p.w, p.h = ps[1], ps[0]
		}
		p.wPt = p.w * p.k
		p.hPt = p.h * p.k
		p.pageBreakTrigger = p.h - p.bMargin
		p.curOrientation = orientation
		p.curPageSize = ps
	}
	if orientation != p.defOrientation || ps != p.defPageSize {
		if p.pageInfo[p.page] == nil {
			p.pageInfo[p.page] = map[string]interface{}{}
		}
		p.pageInfo[p.page]["size"] = [2]float64{p.wPt, p.hPt}
	}
	if rotation != 0 {
		if rotation%90 != 0 {
			p.panicError("incorrect rotation value")
		}
		if p.pageInfo[p.page] == nil {
			p.pageInfo[p.page] = map[string]interface{}{}
		}
		p.pageInfo[p.page]["rotation"] = rotation
	}
	p.curRotation = rotation
}

func (p *G3PDF) endPage() { p.state = 1 }

func (p *G3PDF) out(s string) {
	switch p.state {
	case 2:
		p.pages[p.page] = append(p.pages[p.page], s)
	case 0:
		p.panicError("no page has been added yet")
	case 1:
		p.panicError("invalid call")
	case 3:
		p.panicError("the document is closed")
	}
}

func (p *G3PDF) put(s string) {
	p.buffer.WriteString(s)
	p.buffer.WriteByte('\n')
}

func (p *G3PDF) getOffset() int { return p.buffer.Len() }

func (p *G3PDF) newObj(forced ...int) {
	n := 0
	if len(forced) > 0 {
		n = forced[0]
		p.n = maxInt(p.n, n)
	} else {
		p.n++
		n = p.n
	}
	p.offsets[n] = p.getOffset()
	p.put(strconv.Itoa(n) + " 0 obj")
}

func (p *G3PDF) putStream(data []byte) {
	p.put("stream")
	p.buffer.Write(data)
	p.buffer.WriteByte('\n')
	p.put("endstream")
}

func (p *G3PDF) putStreamObject(data []byte) {
	entries := ""
	if p.compress {
		entries = "/Filter /FlateDecode "
		data = flateCompress(data)
	}
	entries += "/Length " + strconv.Itoa(len(data))
	p.newObj()
	p.put("<<" + entries + ">>")
	p.putStream(data)
	p.put("endobj")
}

func (p *G3PDF) putLinks(n int) {
	for _, pl := range p.pageLinks[n] {
		p.newObj()
		x := toFloat(pl[0])
		y := toFloat(pl[1])
		w := toFloat(pl[2])
		h := toFloat(pl[3])
		rect := sprintf("%.2F %.2F %.2F %.2F", x, y, x+w, y-h)
		s := "<</Type /Annot /Subtype /Link /Rect [" + rect + "] /Border [0 0 0] "
		switch v := pl[4].(type) {
		case string:
			s += "/A <</S /URI /URI " + p.textString(v) + ">>>>"
		default:
			lnk := toInt(v)
			dst := p.links[lnk]
			page := int(dst[0])
			y2 := dst[1]
			hPage := p.hPt
			if pi, ok := p.pageInfo[page]; ok {
				if sz, ok2 := pi["size"].([2]float64); ok2 {
					hPage = sz[1]
				}
			}
			nobj := p.pageInfo[page]["n"]
			s += sprintf("/Dest [%d 0 R /XYZ 0 %.2F null]>>", toInt(nobj), hPage-y2*p.k)
		}
		p.put(s)
		p.put("endobj")
	}
}

func (p *G3PDF) putPage(n int) {
	p.newObj()
	p.put("<</Type /Page")
	p.put("/Parent 1 0 R")
	if pi, ok := p.pageInfo[n]; ok {
		if sz, ok2 := pi["size"].([2]float64); ok2 {
			p.put(sprintf("/MediaBox [0 0 %.2F %.2F]", sz[0], sz[1]))
		}
		if rot, ok2 := pi["rotation"].(int); ok2 {
			p.put("/Rotate " + strconv.Itoa(rot))
		}
	}
	p.put("/Resources 2 0 R")
	if len(p.pageLinks[n]) > 0 {
		s := "/Annots ["
		for _, pl := range p.pageLinks[n] {
			s += strconv.Itoa(toInt(pl[5])) + " 0 R "
		}
		s += "]"
		p.put(s)
	}
	if p.withAlpha {
		p.put("/Group <</Type /Group /S /Transparency /CS /DeviceRGB>>")
	}
	p.put("/Contents " + strconv.Itoa(p.n+1) + " 0 R>>")
	p.put("endobj")

	content := strings.Join(p.pages[n], "\n") + "\n"
	if p.aliasNbPages != "" {
		content = strings.ReplaceAll(content, p.aliasNbPages, strconv.Itoa(p.page))
	}
	p.putStreamObject([]byte(content))
	p.putLinks(n)
}

func (p *G3PDF) putPages() {
	n := p.n
	for i := 1; i <= p.page; i++ {
		if p.pageInfo[i] == nil {
			p.pageInfo[i] = map[string]interface{}{}
		}
		n++
		p.pageInfo[i]["n"] = n
		n++
		for idx := range p.pageLinks[i] {
			n++
			p.pageLinks[i][idx] = append(p.pageLinks[i][idx], n)
		}
	}
	for i := 1; i <= p.page; i++ {
		p.putPage(i)
	}
	p.newObj(1)
	p.put("<</Type /Pages")
	kids := "/Kids ["
	for i := 1; i <= p.page; i++ {
		kids += strconv.Itoa(toInt(p.pageInfo[i]["n"])) + " 0 R "
	}
	kids += "]"
	p.put(kids)
	p.put("/Count " + strconv.Itoa(p.page))
	w, h := p.defPageSize[0], p.defPageSize[1]
	if p.defOrientation != "P" {
		w, h = h, w
	}
	p.put(sprintf("/MediaBox [0 0 %.2F %.2F]", w*p.k, h*p.k))
	p.put(">>")
	p.put("endobj")
}

func (p *G3PDF) putFonts() {
	for k, f := range p.fonts {
		toUnicodeObj := 0
		if len(f.uv) > 0 {
			cmap := p.toUnicodeCMap(f.uv)
			p.putStreamObject([]byte(cmap))
			toUnicodeObj = p.n
		}

		p.newObj()
		f.n = p.n
		p.fonts[k] = f

		p.put("<</Type /Font")
		p.put("/BaseFont /" + f.name)
		p.put("/Subtype /Type1")
		if f.name != "Symbol" && f.name != "ZapfDingbats" {
			p.put("/Encoding /WinAnsiEncoding")
		}
		if toUnicodeObj > 0 {
			p.put("/ToUnicode " + strconv.Itoa(toUnicodeObj) + " 0 R")
		}
		p.put(">>")
		p.put("endobj")
	}
}

func (p *G3PDF) toUnicodeCMap(uv map[int]interface{}) string {
	var ranges strings.Builder
	var chars strings.Builder
	nbr, nbc := 0, 0
	keys := make([]int, 0, len(uv))
	for k := range uv {
		keys = append(keys, k)
	}
	sortInts(keys)
	for _, c := range keys {
		v := uv[c]
		switch vv := v.(type) {
		case pdfUVRange:
			ranges.WriteString(sprintf("<%02X> <%02X> <%04X>\n", c, c+vv.count-1, vv.start))
			nbr++
		case int:
			chars.WriteString(sprintf("<%02X> <%04X>\n", c, vv))
			nbc++
		}
	}
	var b strings.Builder
	b.WriteString("/CIDInit /ProcSet findresource begin\n")
	b.WriteString("12 dict begin\n")
	b.WriteString("begincmap\n")
	b.WriteString("/CIDSystemInfo\n<</Registry (Adobe)\n/Ordering (UCS)\n/Supplement 0\n>> def\n")
	b.WriteString("/CMapName /Adobe-Identity-UCS def\n/CMapType 2 def\n")
	b.WriteString("1 begincodespacerange\n<00> <FF>\nendcodespacerange\n")
	if nbr > 0 {
		b.WriteString(strconv.Itoa(nbr) + " beginbfrange\n")
		b.WriteString(ranges.String())
		b.WriteString("endbfrange\n")
	}
	if nbc > 0 {
		b.WriteString(strconv.Itoa(nbc) + " beginbfchar\n")
		b.WriteString(chars.String())
		b.WriteString("endbfchar\n")
	}
	b.WriteString("endcmap\nCMapName currentdict /CMap defineresource pop\nend\nend")
	return b.String()
}

func (p *G3PDF) putImages() {
	for _, info := range p.images {
		p.putImage(info)
	}
}

func (p *G3PDF) putImage(info *pdfImage) {
	p.newObj()
	info.n = p.n
	p.put("<</Type /XObject")
	p.put("/Subtype /Image")
	p.put("/Width " + strconv.Itoa(info.w))
	p.put("/Height " + strconv.Itoa(info.h))
	p.put("/ColorSpace /" + info.cs)
	p.put("/BitsPerComponent " + strconv.Itoa(info.bpc))
	if info.f != "" {
		p.put("/Filter /" + info.f)
	}
	p.put("/Length " + strconv.Itoa(len(info.data)) + ">>")
	p.putStream(info.data)
	p.put("endobj")
}

func (p *G3PDF) putXObjectDict() {
	for _, image := range p.images {
		p.put("/I" + strconv.Itoa(image.i) + " " + strconv.Itoa(image.n) + " 0 R")
	}
}

func (p *G3PDF) putResourceDict() {
	p.put("/ProcSet [/PDF /Text /ImageB /ImageC /ImageI]")
	p.put("/Font <<")
	for _, f := range p.fonts {
		p.put("/F" + strconv.Itoa(f.i) + " " + strconv.Itoa(f.n) + " 0 R")
	}
	p.put(">>")
	p.put("/XObject <<")
	p.putXObjectDict()
	p.put(">>")
}

func (p *G3PDF) putResources() {
	p.putFonts()
	p.putImages()
	p.newObj(2)
	p.put("<<")
	p.putResourceDict()
	p.put(">>")
	p.put("endobj")
}

func (p *G3PDF) putInfo() {
	date := p.creationDate.Format("20060102150405-0700")
	p.metadata["CreationDate"] = "D:" + date[:len(date)-2] + "'" + date[len(date)-2:] + "'"
	keys := make([]string, 0, len(p.metadata))
	for k := range p.metadata {
		keys = append(keys, k)
	}
	sortStrings(keys)
	for _, k := range keys {
		p.put("/" + k + " " + p.textString(p.metadata[k]))
	}
}

func (p *G3PDF) putCatalog() {
	n := toInt(p.pageInfo[1]["n"])
	p.put("/Type /Catalog")
	p.put("/Pages 1 0 R")
	switch v := p.zoomMode.(type) {
	case string:
		s := strings.ToLower(v)
		switch s {
		case "fullpage":
			p.put("/OpenAction [" + strconv.Itoa(n) + " 0 R /Fit]")
		case "fullwidth":
			p.put("/OpenAction [" + strconv.Itoa(n) + " 0 R /FitH null]")
		case "real":
			p.put("/OpenAction [" + strconv.Itoa(n) + " 0 R /XYZ null null 1]")
		}
	case float64:
		p.put(sprintf("/OpenAction [%d 0 R /XYZ null null %.2F]", n, v/100))
	}
	switch p.layoutMode {
	case "single":
		p.put("/PageLayout /SinglePage")
	case "continuous":
		p.put("/PageLayout /OneColumn")
	case "two":
		p.put("/PageLayout /TwoColumnLeft")
	}
}

func (p *G3PDF) putHeader() { p.put("%PDF-" + p.pdfVersion) }

func (p *G3PDF) putTrailer() {
	p.put("/Size " + strconv.Itoa(p.n+1))
	p.put("/Root " + strconv.Itoa(p.n) + " 0 R")
	p.put("/Info " + strconv.Itoa(p.n-1) + " 0 R")
}

func (p *G3PDF) endDoc() {
	p.creationDate = time.Now()
	p.putHeader()
	p.putPages()
	p.putResources()
	p.newObj()
	p.put("<<")
	p.putInfo()
	p.put(">>")
	p.put("endobj")
	p.newObj()
	p.put("<<")
	p.putCatalog()
	p.put(">>")
	p.put("endobj")
	offset := p.getOffset()
	p.put("xref")
	p.put("0 " + strconv.Itoa(p.n+1))
	p.put("0000000000 65535 f ")
	for i := 1; i <= p.n; i++ {
		p.put(sprintf("%010d 00000 n ", p.offsets[i]))
	}
	p.put("trailer")
	p.put("<<")
	p.putTrailer()
	p.put(">>")
	p.put("startxref")
	p.put(strconv.Itoa(offset))
	p.put("%%EOF")
	p.state = 3
}

func (p *G3PDF) httpEncode(param, value string, isUTF8 bool) string {
	if isASCII(value) {
		return param + "=\"" + value + "\""
	}
	if !isUTF8 {
		value = latin1ToUTF8(value)
	}
	return param + "*=UTF-8''" + urlEncode(value)
}

func (p *G3PDF) escape(s string) string {
	r := strings.ReplaceAll(s, "\\", "\\\\")
	r = strings.ReplaceAll(r, "(", "\\(")
	r = strings.ReplaceAll(r, ")", "\\)")
	r = strings.ReplaceAll(r, "\r", "\\r")
	return r
}

func (p *G3PDF) textString(s string) string {
	if !isASCII(s) {
		s = utf8ToUTF16BEWithBOM(s)
	}
	return "(" + p.escape(s) + ")"
}

func (p *G3PDF) doUnderline(x, y float64, txt string) string {
	if p.currentFont == nil {
		return ""
	}
	w := p.GetStringWidth(txt) + p.ws*float64(strings.Count(txt, " "))
	return sprintf("%.2F %.2F %.2F %.2F re f", x*p.k, (p.h-(y-p.currentFont.up/1000*p.fontSize))*p.k, w*p.k, -p.currentFont.ut/1000*p.fontSizePt)
}

func (p *G3PDF) parseImageFile(file string) *pdfImage {
	f, err := os.Open(file)
	if err != nil {
		p.panicError("can't open image file: " + file)
	}
	defer f.Close()

	cfg, format, err := image.DecodeConfig(f)
	if err != nil {
		p.panicError("missing or incorrect image file: " + file)
	}

	if _, err := f.Seek(0, io.SeekStart); err != nil {
		p.panicError("unable to seek image file")
	}

	switch strings.ToLower(format) {
	case "jpeg":
		data, readErr := io.ReadAll(f)
		if readErr != nil {
			p.panicError("unable to read JPEG image file")
		}
		return &pdfImage{w: cfg.Width, h: cfg.Height, cs: "DeviceRGB", bpc: 8, f: "DCTDecode", data: data}
	default:
		img, _, decodeErr := image.Decode(f)
		if decodeErr != nil {
			p.panicError("unable to decode image file: " + file)
		}

		var encoded bytes.Buffer
		if encodeErr := stdjpeg.Encode(&encoded, img, &stdjpeg.Options{Quality: 90}); encodeErr != nil {
			p.panicError("unable to encode image as JPEG: " + file)
		}

		return &pdfImage{w: cfg.Width, h: cfg.Height, cs: "DeviceRGB", bpc: 8, f: "DCTDecode", data: encoded.Bytes()}
	}
}

func (p *G3PDF) charWidth(c byte) int {
	if p.currentFont == nil {
		return 0
	}
	w := p.currentFont.cw[c]
	if w == 0 {
		return p.currentFont.cw['?']
	}
	return w
}

func (p *G3PDF) loadFontAsset(file string) (*pdfFont, bool) {
	key := strings.ToLower(filepath.Base(file))
	f, ok := p.assetFonts[key]
	if !ok {
		return nil, false
	}
	return f, true
}

func sprintf(format string, args ...interface{}) string { return fmt.Sprintf(format, args...) }

func containsString(list []string, v string) bool {
	for _, x := range list {
		if x == v {
			return true
		}
	}
	return false
}

func maxInt(a, b int) int {
	if a > b {
		return a
	}
	return b
}

func flateCompress(data []byte) []byte {
	var b bytes.Buffer
	w := zlib.NewWriter(&b)
	_, _ = w.Write(data)
	_ = w.Close()
	return b.Bytes()
}

func isASCII(s string) bool {
	for i := 0; i < len(s); i++ {
		if s[i] > 127 {
			return false
		}
	}
	return true
}

func latin1ToUTF8(s string) string {
	var b strings.Builder
	for i := 0; i < len(s); i++ {
		c := s[i]
		if c >= 128 {
			b.WriteByte(0xC0 | (c >> 6))
			b.WriteByte(0x80 | (c & 0x3F))
		} else {
			b.WriteByte(c)
		}
	}
	return b.String()
}

func utf8ToUTF16BEWithBOM(s string) string {
	runes := []rune(s)
	buf := make([]byte, 2, 2+len(runes)*2)
	buf[0] = 0xFE
	buf[1] = 0xFF
	for _, r := range runes {
		if r > 0xFFFF {
			r = '?'
		}
		tmp := make([]byte, 2)
		binary.BigEndian.PutUint16(tmp, uint16(r))
		buf = append(buf, tmp...)
	}
	return string(buf)
}

func urlEncode(s string) string {
	var b strings.Builder
	for i := 0; i < len(s); i++ {
		c := s[i]
		if (c >= 'a' && c <= 'z') || (c >= 'A' && c <= 'Z') || (c >= '0' && c <= '9') || c == '-' || c == '_' || c == '.' || c == '~' {
			b.WriteByte(c)
		} else {
			b.WriteString(sprintf("%%%02X", c))
		}
	}
	return b.String()
}

func sortInts(a []int) {
	for i := 1; i < len(a); i++ {
		j := i
		for j > 0 && a[j-1] > a[j] {
			a[j-1], a[j] = a[j], a[j-1]
			j--
		}
	}
}

func sortStrings(a []string) {
	for i := 1; i < len(a); i++ {
		j := i
		for j > 0 && a[j-1] > a[j] {
			a[j-1], a[j] = a[j], a[j-1]
			j--
		}
	}
}
