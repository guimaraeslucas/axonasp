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
package axonvm

import (
	"bytes"
	"errors"
	"fmt"
	"image"
	"image/color"
	"image/jpeg"
	"image/png"
	"os"
	"path/filepath"
	"strings"

	"github.com/fogleman/gg"
	"golang.org/x/image/font"
)

type G3Image struct {
	vm            *VM
	objectID      int64
	dc            *gg.Context
	lastErr       string
	lastBytes     []byte
	lastMimeType  string
	lastTempFile  string
	lastLoaded    image.Image
	lastFontFace  font.Face
	defaultFormat string
	jpgQuality    int
}

// newG3ImageObject instantiates the G3Image custom functions library.
func (vm *VM) newG3ImageObject() Value {
	obj := &G3Image{
		vm:            vm,
		defaultFormat: "png",
		jpgQuality:    90,
	}
	id := vm.nextDynamicNativeID
	vm.nextDynamicNativeID++
	obj.objectID = id
	vm.g3imageItems[id] = obj
	return Value{Type: VTNativeObject, Num: id}
}

// DispatchPropertyGet acts as a getter.
func (g *G3Image) DispatchPropertyGet(propertyName string) Value {
	switch strings.ToLower(propertyName) {
	case "hascontext":
		return NewBool(g.dc != nil)
	case "width":
		if g.dc == nil {
			return NewInteger(0)
		}
		return NewInteger(int64(g.dc.Width()))
	case "height":
		if g.dc == nil {
			return NewInteger(0)
		}
		return NewInteger(int64(g.dc.Height()))
	case "lasterror":
		return NewString(g.lastErr)
	case "lastmimetype", "contenttype", "mimetype":
		return NewString(g.lastMimeType)
	case "lasttempfile", "tempfile":
		return NewString(g.lastTempFile)
	case "lastbytes", "content":
		if g.lastBytes == nil {
			return Value{Type: VTArray, Arr: NewVBArrayFromValues(0, []Value{})}
		}
		// Convert to VM Array
		arr := make([]Value, len(g.lastBytes))
		for i, b := range g.lastBytes {
			arr[i] = NewInteger(int64(b))
		}
		return Value{Type: VTArray, Arr: NewVBArrayFromValues(0, arr)}
	case "defaultformat":
		return NewString(g.defaultFormat)
	case "jpgquality", "jpegquality":
		return NewInteger(int64(g.jpgQuality))
	// GG Constants
	case "alignleft":
		return NewInteger(int64(gg.AlignLeft))
	case "aligncenter":
		return NewInteger(int64(gg.AlignCenter))
	case "alignright":
		return NewInteger(int64(gg.AlignRight))
	case "fillrulewinding":
		return NewInteger(int64(gg.FillRuleWinding))
	case "fillruleevenodd":
		return NewInteger(int64(gg.FillRuleEvenOdd))
	case "linecapround":
		return NewInteger(int64(gg.LineCapRound))
	case "linecapbutt":
		return NewInteger(int64(gg.LineCapButt))
	case "linecapsquare":
		return NewInteger(int64(gg.LineCapSquare))
	case "linejoinround":
		return NewInteger(int64(gg.LineJoinRound))
	case "linejoinbevel":
		return NewInteger(int64(gg.LineJoinBevel))
	}
	return g.DispatchMethod(propertyName, nil)
}

// DispatchPropertySet acts as a setter.
func (g *G3Image) DispatchPropertySet(propertyName string, args []Value) bool {
	if len(args) == 0 {
		return false
	}
	val := args[0]
	switch strings.ToLower(propertyName) {
	case "defaultformat":
		format := strings.ToLower(strings.TrimSpace(val.String()))
		if format == "png" || format == "jpg" || format == "jpeg" {
			g.defaultFormat = format
		}
		return true
	case "jpgquality", "jpegquality":
		q := int(g.vm.asInt(val))
		if q < 1 {
			q = 1
		}
		if q > 100 {
			q = 100
		}
		g.jpgQuality = q
		return true
	}
	return false
}

// DispatchMethod provides O(1) string matching resolution.
func (g *G3Image) DispatchMethod(methodName string, args []Value) Value {
	method := strings.ToLower(strings.TrimSpace(methodName))

	switch method {
	case "close", "dispose", "release", "destroy", "reset", "clearcontext", "clearimage":
		g.closeAndDetach()
		return NewBool(true)

	case "new", "newcontext", "create", "createcontext", "init":
		if len(args) < 2 {
			g.setError("newcontext requires width and height")
			return NewEmpty()
		}
		g.releaseResources(false)
		w := int(g.vm.asInt(args[0]))
		h := int(g.vm.asInt(args[1]))
		if w <= 0 || h <= 0 {
			g.setError("newcontext requires positive dimensions")
			return NewEmpty()
		}
		g.dc = gg.NewContext(w, h)
		g.clearError()
		return NewBool(true)

	case "loadimage", "load":
		if len(args) < 1 {
			g.setError("loadimage requires path")
			return NewEmpty()
		}
		im, err := g.loadImageFromPath(args[0].String(), "")
		if err != nil {
			g.setError(err.Error())
			return NewEmpty()
		}
		g.lastLoaded = im
		g.clearError()
		return NewBool(true)

	case "loadpng":
		if len(args) < 1 {
			g.setError("loadpng requires path")
			return NewEmpty()
		}
		im, err := g.loadImageFromPath(args[0].String(), "png")
		if err != nil {
			g.setError(err.Error())
			return NewEmpty()
		}
		g.lastLoaded = im
		g.clearError()
		return NewBool(true)

	case "loadjpg", "loadjpeg":
		if len(args) < 1 {
			g.setError("loadjpg requires path")
			return NewEmpty()
		}
		im, err := g.loadImageFromPath(args[0].String(), "jpg")
		if err != nil {
			g.setError(err.Error())
			return NewEmpty()
		}
		g.lastLoaded = im
		g.clearError()
		return NewBool(true)

	case "newcontextforimage", "contextforimage", "useimage", "setimage":
		if g.lastLoaded == nil {
			g.setError("no image loaded to create context for")
			return NewBool(false)
		}
		g.dc = gg.NewContextForImage(g.lastLoaded)
		g.clearError()
		return NewBool(true)

	case "savepng":
		if g.dc == nil {
			g.setError("no active context")
			return NewBool(false)
		}
		if len(args) < 1 {
			g.setError("savepng requires path")
			return NewBool(false)
		}
		err := g.savePNGToPath(args[0].String())
		if err != nil {
			g.setError(err.Error())
			return NewBool(false)
		}
		g.clearError()
		return NewBool(true)

	case "savejpg", "savejpeg":
		if g.dc == nil {
			g.setError("no active context")
			return NewBool(false)
		}
		if len(args) < 1 {
			g.setError("savejpg requires path")
			return NewBool(false)
		}
		quality := g.jpgQuality
		if len(args) > 1 {
			quality = int(g.vm.asInt(args[1]))
		}
		err := g.saveJPGToPath(args[0].String(), quality)
		if err != nil {
			g.setError(err.Error())
			return NewBool(false)
		}
		g.clearError()
		return NewBool(true)

	// Context specific rendering commands without interfaces
	case "sethexcolor":
		if len(args) < 1 || g.dc == nil {
			return NewEmpty()
		}
		g.dc.SetHexColor(args[0].String())
		return NewEmpty()

	case "setcolor":
		if len(args) < 1 || g.dc == nil {
			return NewEmpty()
		}
		c, err := parseColorString(args[0].String())
		if err == nil {
			g.dc.SetColor(c)
		}
		return NewEmpty()

	case "clear":
		if g.dc != nil {
			g.dc.Clear()
		}
		return NewEmpty()

	case "setlinewidth":
		if len(args) < 1 || g.dc == nil {
			return NewEmpty()
		}
		g.dc.SetLineWidth(g.vm.asFloat(args[0]))
		return NewEmpty()

	case "drawline":
		if len(args) < 4 || g.dc == nil {
			return NewEmpty()
		}
		g.dc.DrawLine(g.vm.asFloat(args[0]), g.vm.asFloat(args[1]), g.vm.asFloat(args[2]), g.vm.asFloat(args[3]))
		return NewEmpty()

	case "drawrectangle":
		if len(args) < 4 || g.dc == nil {
			return NewEmpty()
		}
		g.dc.DrawRectangle(g.vm.asFloat(args[0]), g.vm.asFloat(args[1]), g.vm.asFloat(args[2]), g.vm.asFloat(args[3]))
		return NewEmpty()

	case "drawcircle":
		if len(args) < 3 || g.dc == nil {
			return NewEmpty()
		}
		g.dc.DrawCircle(g.vm.asFloat(args[0]), g.vm.asFloat(args[1]), g.vm.asFloat(args[2]))
		return NewEmpty()

	case "drawellipse":
		if len(args) < 4 || g.dc == nil {
			return NewEmpty()
		}
		g.dc.DrawEllipse(g.vm.asFloat(args[0]), g.vm.asFloat(args[1]), g.vm.asFloat(args[2]), g.vm.asFloat(args[3]))
		return NewEmpty()

	case "stroke":
		if g.dc != nil {
			g.dc.Stroke()
		}
		return NewEmpty()

	case "fill":
		if g.dc != nil {
			g.dc.Fill()
		}
		return NewEmpty()

	case "fillpreserve":
		if g.dc != nil {
			g.dc.FillPreserve()
		}
		return NewEmpty()

	case "strokepreserve":
		if g.dc != nil {
			g.dc.StrokePreserve()
		}
		return NewEmpty()

	case "loadfontface":
		if len(args) < 2 {
			g.setError("loadfontface requires path and points")
			return NewBool(false)
		}
		fontPath, err := g.resolveRootPath(args[0].String())
		if err != nil {
			g.setError(err.Error())
			return NewBool(false)
		}
		f, err := gg.LoadFontFace(fontPath, g.vm.asFloat(args[1]))
		if err != nil {
			g.setError(err.Error())
			return NewBool(false)
		}
		g.lastFontFace = f
		if g.dc != nil {
			g.dc.SetFontFace(f)
		}
		g.clearError()
		return NewBool(true)

	case "drawstring":
		if len(args) < 3 || g.dc == nil {
			return NewEmpty()
		}
		g.dc.DrawString(args[0].String(), g.vm.asFloat(args[1]), g.vm.asFloat(args[2]))
		return NewEmpty()

	case "drawstringanchored":
		if len(args) < 5 || g.dc == nil {
			return NewEmpty()
		}
		g.dc.DrawStringAnchored(args[0].String(), g.vm.asFloat(args[1]), g.vm.asFloat(args[2]), g.vm.asFloat(args[3]), g.vm.asFloat(args[4]))
		return NewEmpty()

	case "measurestring":
		if len(args) < 1 || g.dc == nil {
			return NewEmpty()
		}
		w, h := g.dc.MeasureString(args[0].String())
		arr := make([]Value, 2)
		arr[0] = NewDouble(w)
		arr[1] = NewDouble(h)
		return Value{Type: VTArray, Arr: NewVBArrayFromValues(0, arr)}

	case "drawimage":
		if len(args) < 2 || g.dc == nil || g.lastLoaded == nil {
			return NewEmpty()
		}
		g.dc.DrawImage(g.lastLoaded, int(g.vm.asInt(args[0])), int(g.vm.asInt(args[1])))
		return NewEmpty()

	case "renderviatemp", "renderbytemp", "getcontentviatemp", "rendertemp":
		format := g.defaultFormat
		quality := g.jpgQuality
		if len(args) > 0 {
			format = strings.ToLower(strings.TrimSpace(args[0].String()))
		}
		if len(args) > 1 {
			quality = int(g.vm.asInt(args[1]))
		}
		data, err := g.renderViaTemp(format, quality)
		if err != nil {
			g.setError(err.Error())
			return NewEmpty()
		}
		g.clearError()
		arr := make([]Value, len(data))
		for i, b := range data {
			arr[i] = NewInteger(int64(b))
		}
		return Value{Type: VTArray, Arr: NewVBArrayFromValues(0, arr)}
	}

	return NewEmpty()
}

// closeAndDetach releases all image buffers and unregisters this object from VM maps.
func (g *G3Image) closeAndDetach() {
	if g == nil {
		return
	}
	g.releaseResources(true)
	if g.vm == nil {
		g.objectID = 0
		return
	}
	if g.objectID != 0 {
		delete(g.vm.g3imageItems, g.objectID)
		delete(g.vm.nativeObjectProxies, g.objectID)
		g.objectID = 0
		return
	}
	for id, item := range g.vm.g3imageItems {
		if item != g {
			continue
		}
		delete(g.vm.g3imageItems, id)
		delete(g.vm.nativeObjectProxies, id)
		break
	}
}

func (g *G3Image) releaseResources(clearError bool) {
	g.dc = nil
	g.lastBytes = nil
	g.lastMimeType = ""
	g.lastTempFile = ""
	g.lastLoaded = nil
	g.lastFontFace = nil
	if clearError {
		g.lastErr = ""
	}

}

// cleanupG3ImageResources releases all image contexts owned by one VM request.
func (vm *VM) cleanupG3ImageResources() {
	if vm == nil || len(vm.g3imageItems) == 0 {
		return
	}
	for id, item := range vm.g3imageItems {
		if item != nil {
			item.releaseResources(false)
			item.objectID = 0
		}
		delete(vm.g3imageItems, id)
		delete(vm.nativeObjectProxies, id)
	}
}

func (g *G3Image) renderPNGBytes() ([]byte, error) {
	if g.dc == nil {
		return nil, errors.New("no active context")
	}
	var buf bytes.Buffer
	if err := g.dc.EncodePNG(&buf); err != nil {
		return nil, err
	}
	g.lastMimeType = "image/png"
	g.lastTempFile = ""
	g.lastBytes = buf.Bytes()
	return g.lastBytes, nil
}

func (g *G3Image) renderJPGBytes(quality int) ([]byte, error) {
	if g.dc == nil {
		return nil, errors.New("no active context")
	}
	if quality < 1 || quality > 100 {
		quality = g.jpgQuality
	}
	var buf bytes.Buffer
	if err := jpeg.Encode(&buf, g.dc.Image(), &jpeg.Options{Quality: quality}); err != nil {
		return nil, err
	}
	g.lastMimeType = "image/jpeg"
	g.lastTempFile = ""
	g.lastBytes = buf.Bytes()
	return g.lastBytes, nil
}

func (g *G3Image) renderViaTemp(format string, quality int) ([]byte, error) {
	if g.dc == nil {
		return nil, errors.New("no active context")
	}
	format = normalizeImageFormat(format)
	if quality < 1 || quality > 100 {
		quality = g.jpgQuality
	}

	tempDir, err := executableTempImagesDir()
	if err != nil {
		return nil, err
	}

	ext := ".png"
	mime := "image/png"
	if format == "jpg" {
		ext = ".jpg"
		mime = "image/jpeg"
	}

	tmpFile, err := os.CreateTemp(tempDir, "axonasp_img_*"+ext)
	if err != nil {
		return nil, err
	}
	tmpPath := tmpFile.Name()
	_ = tmpFile.Close()

	defer func() {
		_ = os.Remove(tmpPath)
	}()

	var encodeErr error
	if format == "jpg" {
		encodeErr = g.saveJPGAbsolute(tmpPath, quality)
	} else {
		encodeErr = g.savePNGAbsolute(tmpPath)
	}
	if encodeErr != nil {
		return nil, encodeErr
	}

	data, err := os.ReadFile(tmpPath)
	if err != nil {
		return nil, err
	}

	g.lastBytes = data
	g.lastMimeType = mime
	g.lastTempFile = tmpPath
	return data, nil
}

func (g *G3Image) savePNGToPath(path string) error {
	fullPath, err := g.resolveRootPath(path)
	if err != nil {
		return err
	}
	return g.savePNGAbsolute(fullPath)
}

func (g *G3Image) saveJPGToPath(path string, quality int) error {
	fullPath, err := g.resolveRootPath(path)
	if err != nil {
		return err
	}
	return g.saveJPGAbsolute(fullPath, quality)
}

func (g *G3Image) savePNGAbsolute(path string) error {
	if g.dc == nil {
		return errors.New("no active context")
	}
	if err := os.MkdirAll(filepath.Dir(path), 0755); err != nil {
		return err
	}
	file, err := os.Create(path)
	if err != nil {
		return err
	}
	defer file.Close()
	if err := png.Encode(file, g.dc.Image()); err != nil {
		return err
	}
	g.lastMimeType = "image/png"
	g.lastTempFile = path
	return nil
}

func (g *G3Image) saveJPGAbsolute(path string, quality int) error {
	if g.dc == nil {
		return errors.New("no active context")
	}
	if quality < 1 || quality > 100 {
		quality = g.jpgQuality
	}
	if err := os.MkdirAll(filepath.Dir(path), 0755); err != nil {
		return err
	}
	file, err := os.Create(path)
	if err != nil {
		return err
	}
	defer file.Close()
	if err := jpeg.Encode(file, g.dc.Image(), &jpeg.Options{Quality: quality}); err != nil {
		return err
	}
	g.lastMimeType = "image/jpeg"
	g.lastTempFile = path
	return nil
}

func (g *G3Image) resolveRootPath(rel string) (string, error) {
	rel = strings.TrimSpace(rel)
	if rel == "" {
		return "", errors.New("path is required")
	}

	if g.vm.host == nil || g.vm.host.Server() == nil {
		return filepath.Abs(rel)
	}

	mapped := g.vm.host.Server().MapPath(rel)
	if mapped == "" {
		return "", errors.New("invalid mapped path")
	}
	return filepath.Abs(mapped)
}

func executableTempImagesDir() (string, error) {
	execPath, err := os.Executable()
	if err != nil {
		return "", err
	}
	dir := filepath.Join(filepath.Dir(execPath), "temp", "images")
	if err := os.MkdirAll(dir, 0755); err != nil {
		return "", err
	}
	return dir, nil
}

func (g *G3Image) loadImageFromPath(path string, kind string) (image.Image, error) {
	fullPath, err := g.resolveRootPath(path)
	if err != nil {
		return nil, err
	}

	switch strings.ToLower(kind) {
	case "png":
		return gg.LoadPNG(fullPath)
	case "jpg", "jpeg":
		return gg.LoadJPG(fullPath)
	default:
		return gg.LoadImage(fullPath)
	}
}

func (g *G3Image) clearError() {
	g.lastErr = ""
}

func (g *G3Image) setError(err string) {
	g.lastErr = err
}

func normalizeImageFormat(format string) string {
	format = strings.ToLower(strings.TrimSpace(format))
	if format == "jpeg" {
		return "jpg"
	}
	if format != "jpg" {
		return "png"
	}
	return format
}

func parseColorString(s string) (color.Color, error) {
	v := strings.TrimSpace(strings.ToLower(s))
	v = strings.TrimPrefix(v, "#")

	if len(v) == 3 {
		r := strings.Repeat(string(v[0]), 2)
		g := strings.Repeat(string(v[1]), 2)
		b := strings.Repeat(string(v[2]), 2)
		return parseHexRGBA(r + g + b + "ff")
	}
	if len(v) == 4 {
		r := strings.Repeat(string(v[0]), 2)
		g := strings.Repeat(string(v[1]), 2)
		b := strings.Repeat(string(v[2]), 2)
		a := strings.Repeat(string(v[3]), 2)
		return parseHexRGBA(r + g + b + a)
	}
	if len(v) == 6 {
		return parseHexRGBA(v + "ff")
	}
	if len(v) == 8 {
		return parseHexRGBA(v)
	}

	parts := strings.Split(v, ",")
	if len(parts) == 3 || len(parts) == 4 {
		r := uint8(0) // Default fallback
		g := uint8(0)
		b := uint8(0)
		a := uint8(255)

		fmt.Sscanf(strings.TrimSpace(parts[0]), "%d", &r)
		fmt.Sscanf(strings.TrimSpace(parts[1]), "%d", &g)
		fmt.Sscanf(strings.TrimSpace(parts[2]), "%d", &b)

		if len(parts) == 4 {
			fmt.Sscanf(strings.TrimSpace(parts[3]), "%d", &a)
		}
		return color.NRGBA{R: r, G: g, B: b, A: a}, nil
	}

	return nil, fmt.Errorf("invalid color string: %s", s)
}

func parseHexRGBA(hex string) (color.Color, error) {
	if len(hex) != 8 {
		return nil, errors.New("hex color must have 8 digits")
	}
	var rgba [4]uint8
	for i := 0; i < 4; i++ {
		var b uint8
		_, err := fmt.Sscanf(hex[i*2:i*2+2], "%02x", &b)
		if err != nil {
			return nil, err
		}
		rgba[i] = b
	}
	return color.NRGBA{R: rgba[0], G: rgba[1], B: rgba[2], A: rgba[3]}, nil
}
