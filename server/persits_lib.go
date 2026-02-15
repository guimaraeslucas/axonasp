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
 */
package server

import (
	"bytes"
	"fmt"
	"image"
	_ "image/gif"
	"image/jpeg"
	"image/png"
	"os"
	"strings"

	xdraw "golang.org/x/image/draw"
)

// PersitsJpeg implements the Persits.Jpeg (AspJpeg) COM object interface.
// It handles image loading from binary data or files, resizing, and
// serving binary image output to the ASP Response.
type PersitsJpeg struct {
	ctx            *ExecutionContext
	loadedImage    image.Image
	originalWidth  int
	originalHeight int
	targetWidth    int
	targetHeight   int
	pngOutput      bool
	jpegQuality    int
	lastBinary     []byte
	dirty          bool
	regKey         string
}

func NewPersitsJpeg(ctx *ExecutionContext) *PersitsJpeg {
	return &PersitsJpeg{
		ctx:         ctx,
		jpegQuality: 85,
	}
}

func (p *PersitsJpeg) GetProperty(name string) interface{} {
	switch strings.ToLower(name) {
	case "originalwidth":
		return p.originalWidth
	case "originalheight":
		return p.originalHeight
	case "width":
		if p.targetWidth > 0 {
			return p.targetWidth
		}
		return p.originalWidth
	case "height":
		if p.targetHeight > 0 {
			return p.targetHeight
		}
		return p.originalHeight
	case "binary":
		return p.getBinaryOutput()
	case "pngoutput":
		return p.pngOutput
	case "regkey":
		return p.regKey
	case "quality":
		return p.jpegQuality
	}
	return nil
}

func (p *PersitsJpeg) SetProperty(name string, value interface{}) {
	switch strings.ToLower(name) {
	case "width":
		p.targetWidth = toInt(value)
		p.dirty = true
	case "height":
		p.targetHeight = toInt(value)
		p.dirty = true
	case "pngoutput":
		p.pngOutput = toBool(value)
		p.dirty = true
	case "regkey":
		p.regKey = fmt.Sprintf("%v", value)
	case "quality":
		q := toInt(value)
		if q > 0 && q <= 100 {
			p.jpegQuality = q
			p.dirty = true
		}
	}
}

func (p *PersitsJpeg) CallMethod(name string, args ...interface{}) interface{} {
	switch strings.ToLower(name) {
	case "open":
		if len(args) < 1 {
			return nil
		}
		path := fmt.Sprintf("%v", args[0])
		return p.openFile(path)
	case "openbinary":
		if len(args) < 1 {
			return nil
		}
		return p.openBinary(args[0])
	case "torgb":
		return nil
	case "sendbinary":
		p.sendBinary()
		return nil
	case "save", "savebinary":
		if len(args) >= 1 {
			path := fmt.Sprintf("%v", args[0])
			p.saveToFile(path)
		}
		return nil
	case "close":
		return nil
	}
	return nil
}

func (p *PersitsJpeg) openFile(path string) interface{} {
	resolved := path
	if p.ctx != nil {
		resolved = p.ctx.Server_MapPath(path)
	}
	f, err := os.Open(resolved)
	if err != nil {
		return nil
	}
	defer f.Close()

	img, _, err := image.Decode(f)
	if err != nil {
		return nil
	}

	p.setImage(img)
	return nil
}

func (p *PersitsJpeg) openBinary(data interface{}) interface{} {
	// Unwrap Field objects (from rs("field").value or rs("field"))
	if f, ok := data.(*Field); ok {
		data = f.Value
	}

	var raw []byte
	switch v := data.(type) {
	case []byte:
		raw = v
	case string:
		raw = []byte(v)
	default:
		return nil
	}

	if len(raw) == 0 {
		return nil
	}

	img, _, err := image.Decode(bytes.NewReader(raw))
	if err != nil {
		return nil
	}

	p.setImage(img)
	return nil
}

func (p *PersitsJpeg) setImage(img image.Image) {
	p.loadedImage = img
	bounds := img.Bounds()
	p.originalWidth = bounds.Dx()
	p.originalHeight = bounds.Dy()
	p.targetWidth = 0
	p.targetHeight = 0
	p.lastBinary = nil
	p.dirty = true
}

func (p *PersitsJpeg) getBinaryOutput() []byte {
	if p.loadedImage == nil {
		return nil
	}

	if !p.dirty && p.lastBinary != nil {
		return p.lastBinary
	}

	targetW := p.originalWidth
	targetH := p.originalHeight
	if p.targetWidth > 0 {
		targetW = p.targetWidth
	}
	if p.targetHeight > 0 {
		targetH = p.targetHeight
	}

	var src image.Image = p.loadedImage

	if targetW != p.originalWidth || targetH != p.originalHeight {
		dst := image.NewNRGBA(image.Rect(0, 0, targetW, targetH))
		xdraw.CatmullRom.Scale(dst, dst.Bounds(), src, src.Bounds(), xdraw.Over, nil)
		src = dst
	}

	var buf bytes.Buffer
	if p.pngOutput {
		if err := png.Encode(&buf, src); err != nil {
			return nil
		}
	} else {
		if err := jpeg.Encode(&buf, src, &jpeg.Options{Quality: p.jpegQuality}); err != nil {
			return nil
		}
	}

	p.lastBinary = buf.Bytes()
	p.dirty = false
	return p.lastBinary
}

func (p *PersitsJpeg) sendBinary() {
	if p.ctx == nil || p.ctx.Response == nil {
		return
	}

	data := p.getBinaryOutput()
	if data == nil {
		return
	}

	if p.pngOutput {
		p.ctx.Response.SetContentType("image/png")
	} else {
		p.ctx.Response.SetContentType("image/jpeg")
	}

	p.ctx.Response.BinaryWrite(data)
}

func (p *PersitsJpeg) saveToFile(path string) {
	data := p.getBinaryOutput()
	if data == nil {
		return
	}
	resolved := path
	if p.ctx != nil {
		resolved = p.ctx.Server_MapPath(path)
	}
	os.WriteFile(resolved, data, 0644)
}

// --- PersitsPDF stub ---

// PersitsPDF is a stub implementation of the Persits.PDF (AspPDF) COM object.
// PDF-to-image conversion is not feasible in pure Go without external tools.
// This stub prevents errors when ASP code creates the object.
type PersitsPDF struct {
	ctx *ExecutionContext
}

func NewPersitsPDF(ctx *ExecutionContext) *PersitsPDF {
	return &PersitsPDF{ctx: ctx}
}

func (p *PersitsPDF) GetProperty(name string) interface{} {
	switch strings.ToLower(name) {
	case "pagecount":
		return 0
	case "pages":
		return &PersitsPDFPages{}
	}
	return nil
}

func (p *PersitsPDF) SetProperty(name string, value interface{}) {}

func (p *PersitsPDF) CallMethod(name string, args ...interface{}) interface{} {
	switch strings.ToLower(name) {
	case "opendocumentbinary", "opendocument", "open":
		return nil
	case "savehttp", "save":
		return nil
	case "close":
		return nil
	}
	return nil
}

// PersitsPDFPages is a stub for the Pages collection.
type PersitsPDFPages struct{}

func (pp *PersitsPDFPages) GetProperty(name string) interface{} {
	switch strings.ToLower(name) {
	case "count":
		return 0
	}
	return nil
}

func (pp *PersitsPDFPages) SetProperty(name string, value interface{}) {}

func (pp *PersitsPDFPages) CallMethod(name string, args ...interface{}) interface{} {
	switch strings.ToLower(name) {
	case "item":
		return &PersitsPDFPage{}
	}
	return nil
}

// PersitsPDFPage is a stub for a single PDF page.
type PersitsPDFPage struct{}

func (pg *PersitsPDFPage) GetProperty(name string) interface{} { return nil }
func (pg *PersitsPDFPage) SetProperty(name string, value interface{}) {}
func (pg *PersitsPDFPage) CallMethod(name string, args ...interface{}) interface{} {
	switch strings.ToLower(name) {
	case "toimage":
		return nil
	}
	return nil
}
