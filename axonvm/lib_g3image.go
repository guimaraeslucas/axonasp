//go:build !wasm && !lib_g3image_disabled

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
	"image/draw"
	"image/jpeg"
	"image/png"
	"os"
	"path/filepath"
	"strings"

	"g3pix.com.br/axonasp/vbscript"
	"github.com/fogleman/gg"
	xdraw "golang.org/x/image/draw"
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

	// Persits compatibility fields
	isPersits        bool
	originalWidth    int
	originalHeight   int
	interpolationVal int // 1=Nearest, 2=Bilinear, 3=Bicubic

	// Sub-object values (lazy initialized)
	canvasVal Value
	fontVal   Value
	penVal    Value
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

// newPersitsJpegObject instantiates the Persits.Jpeg compatible G3Image library.
func (vm *VM) newPersitsJpegObject() Value {
	val := vm.newG3ImageObject()
	obj := vm.g3imageItems[val.Num]
	obj.isPersits = true
	obj.defaultFormat = "jpg"
	obj.jpgQuality = 90
	obj.interpolationVal = 2 // Bilinear default
	return val
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
	case "jpgquality", "jpegquality", "quality":
		return NewInteger(int64(g.jpgQuality))
	case "originalwidth":
		return NewInteger(int64(g.originalWidth))
	case "originalheight":
		return NewInteger(int64(g.originalHeight))
	case "interpolation":
		return NewInteger(int64(g.interpolationVal))
	case "canvas":
		return g.getCanvas()
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
	case "jpgquality", "jpegquality", "quality":
		q := min(max(int(g.vm.asInt(val)), 0), 100)
		g.jpgQuality = q
		return true
	case "interpolation":
		g.interpolationVal = int(g.vm.asInt(val))
		return true
	case "width":
		w := int(g.vm.asInt(val))
		if w > 0 {
			h := 0
			if g.dc != nil {
				h = g.dc.Height()
			}
			if h > 0 {
				g.resizeImage(w, h)
			}
		}
		return true
	case "height":
		h := int(g.vm.asInt(val))
		if h > 0 {
			w := 0
			if g.dc != nil {
				w = g.dc.Width()
			}
			if w > 0 {
				g.resizeImage(w, h)
			}
		}
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
		if g.isPersits {
			if len(args) < 2 {
				g.vm.raise(vbscript.InternalError, NewAxonASPError(ErrG3ImageInvalidArgCount, nil, AxonASPErrorMessages[ErrG3ImageInvalidArgCount], "axonvm/lib_g3image.go", 0).Error())
				return NewEmpty()
			}
			w := int(g.vm.asInt(args[0]))
			h := int(g.vm.asInt(args[1]))
			if w <= 0 || h <= 0 {
				g.vm.raise(vbscript.InternalError, NewAxonASPError(ErrG3ImageInvalidDimension, nil, AxonASPErrorMessages[ErrG3ImageInvalidDimension], "axonvm/lib_g3image.go", 0).Error())
				return NewEmpty()
			}
			g.releaseResources(false)
			g.dc = gg.NewContext(w, h)

			c := color.Color(color.White)
			if len(args) >= 3 {
				if parsed, err := g.parseColorVal(args[2]); err == nil {
					c = parsed
				}
			}
			g.dc.SetColor(c)
			g.dc.Clear()
			g.originalWidth = w
			g.originalHeight = h
			g.clearError()
			return NewBool(true)
		} else {
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
		}

	case "open":
		if len(args) < 1 {
			g.vm.raise(vbscript.InternalError, NewAxonASPError(ErrG3ImageInvalidArgCount, nil, AxonASPErrorMessages[ErrG3ImageInvalidArgCount], "axonvm/lib_g3image.go", 0).Error())
			return NewEmpty()
		}
		im, err := g.loadImageFromPath(args[0].String(), "")
		if err != nil {
			g.setError(err.Error())
			g.vm.raise(vbscript.InternalError, NewAxonASPError(ErrG3ImageLoadFailed, err, err.Error(), "axonvm/lib_g3image.go", 0).Error())
			return NewEmpty()
		}
		g.lastLoaded = im
		g.originalWidth = im.Bounds().Dx()
		g.originalHeight = im.Bounds().Dy()
		g.dc = gg.NewContextForImage(im)
		g.clearError()
		return NewBool(true)

	case "save":
		if len(args) < 1 {
			g.vm.raise(vbscript.InternalError, NewAxonASPError(ErrG3ImageInvalidArgCount, nil, AxonASPErrorMessages[ErrG3ImageInvalidArgCount], "axonvm/lib_g3image.go", 0).Error())
			return NewEmpty()
		}
		if g.dc == nil {
			g.vm.raise(vbscript.InternalError, NewAxonASPError(ErrG3ImageNotInitialized, nil, AxonASPErrorMessages[ErrG3ImageNotInitialized], "axonvm/lib_g3image.go", 0).Error())
			return NewEmpty()
		}
		path := args[0].String()
		ext := strings.ToLower(filepath.Ext(path))
		var err error
		if ext == ".png" {
			err = g.savePNGToPath(path)
		} else {
			err = g.saveJPGToPath(path, g.jpgQuality)
		}
		if err != nil {
			g.setError(err.Error())
			g.vm.raise(vbscript.InternalError, NewAxonASPError(ErrG3ImageSaveFailed, err, err.Error(), "axonvm/lib_g3image.go", 0).Error())
			return NewEmpty()
		}
		return NewBool(true)

	case "sendbinary":
		if g.dc == nil {
			g.vm.raise(vbscript.InternalError, NewAxonASPError(ErrG3ImageNotInitialized, nil, AxonASPErrorMessages[ErrG3ImageNotInitialized], "axonvm/lib_g3image.go", 0).Error())
			return NewEmpty()
		}
		var buf bytes.Buffer
		err := jpeg.Encode(&buf, g.dc.Image(), &jpeg.Options{Quality: g.jpgQuality})
		if err != nil {
			g.setError(err.Error())
			g.vm.raise(vbscript.InternalError, NewAxonASPError(ErrG3ImageSaveFailed, err, err.Error(), "axonvm/lib_g3image.go", 0).Error())
			return NewEmpty()
		}
		g.lastBytes = buf.Bytes()
		g.lastMimeType = "image/jpeg"

		arr := make([]Value, len(g.lastBytes))
		for i, b := range g.lastBytes {
			arr[i] = NewInteger(int64(b))
		}
		return Value{Type: VTArray, Arr: NewVBArrayFromValues(0, arr)}

	case "crop":
		if len(args) < 4 {
			g.vm.raise(vbscript.InternalError, NewAxonASPError(ErrG3ImageInvalidArgCount, nil, AxonASPErrorMessages[ErrG3ImageInvalidArgCount], "axonvm/lib_g3image.go", 0).Error())
			return NewEmpty()
		}
		if g.dc == nil {
			g.vm.raise(vbscript.InternalError, NewAxonASPError(ErrG3ImageNotInitialized, nil, AxonASPErrorMessages[ErrG3ImageNotInitialized], "axonvm/lib_g3image.go", 0).Error())
			return NewEmpty()
		}
		x0 := int(g.vm.asInt(args[0]))
		y0 := int(g.vm.asInt(args[1]))
		x1 := int(g.vm.asInt(args[2]))
		y1 := int(g.vm.asInt(args[3]))
		err := g.Crop(x0, y0, x1, y1)
		if err != nil {
			g.setError(err.Error())
			g.vm.raise(vbscript.InternalError, NewAxonASPError(ErrInvalidProcedureCallOrArg, err, err.Error(), "axonvm/lib_g3image.go", 0).Error())
			return NewEmpty()
		}
		return NewBool(true)

	case "sharpen":
		if len(args) < 2 || g.dc == nil {
			g.vm.raise(vbscript.InternalError, NewAxonASPError(ErrG3ImageInvalidArgCount, nil, AxonASPErrorMessages[ErrG3ImageInvalidArgCount], "axonvm/lib_g3image.go", 0).Error())
			return NewEmpty()
		}
		radius := g.vm.asFloat(args[0])
		amount := g.vm.asFloat(args[1])
		g.Sharpen(radius, amount)
		return NewEmpty()

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
		if g.isPersits {
			if len(args) < 3 || g.dc == nil {
				g.vm.raise(vbscript.InternalError, NewAxonASPError(ErrG3ImageInvalidArgCount, nil, AxonASPErrorMessages[ErrG3ImageInvalidArgCount], "axonvm/lib_g3image.go", 0).Error())
				return NewEmpty()
			}
			x := int(g.vm.asInt(args[0]))
			y := int(g.vm.asInt(args[1]))
			otherVal := args[2]
			if otherVal.Type != VTNativeObject {
				g.vm.raise(vbscript.InternalError, NewAxonASPError(ErrInvalidProcedureCallOrArg, nil, "DrawImage third argument must be a Jpeg object", "axonvm/lib_g3image.go", 0).Error())
				return NewEmpty()
			}
			otherImg, exists := g.vm.g3imageItems[otherVal.Num]
			if !exists || otherImg == nil || otherImg.dc == nil {
				g.vm.raise(vbscript.InternalError, NewAxonASPError(ErrInvalidProcedureCallOrArg, nil, "DrawImage third argument is an invalid or uninitialized Jpeg object", "axonvm/lib_g3image.go", 0).Error())
				return NewEmpty()
			}
			g.dc.DrawImage(otherImg.dc.Image(), x, y)
			return NewEmpty()
		} else {
			if len(args) < 2 || g.dc == nil || g.lastLoaded == nil {
				return NewEmpty()
			}
			g.dc.DrawImage(g.lastLoaded, int(g.vm.asInt(args[0])), int(g.vm.asInt(args[1])))
			return NewEmpty()
		}

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
	if vm == nil {
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
	for id := range vm.g3imageCanvasItems {
		delete(vm.g3imageCanvasItems, id)
	}
	for id := range vm.g3imageFontItems {
		delete(vm.g3imageFontItems, id)
	}
	for id := range vm.g3imagePenItems {
		delete(vm.g3imagePenItems, id)
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

	if filepath.IsAbs(rel) {
		return filepath.Abs(rel)
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
	for i := range 4 {
		var b uint8
		_, err := fmt.Sscanf(hex[i*2:i*2+2], "%02x", &b)
		if err != nil {
			return nil, err
		}
		rgba[i] = b
	}
	return color.NRGBA{R: rgba[0], G: rgba[1], B: rgba[2], A: rgba[3]}, nil
}

func (g *G3Image) parseColorVal(val Value) (color.Color, error) {
	if val.Type == VTString {
		return parseColorString(val.String())
	}
	n := g.vm.asInt(val)
	// BGR color format standard for classic ASP RGB() values
	r := uint8(n & 0xFF)
	gComp := uint8((n >> 8) & 0xFF)
	b := uint8((n >> 16) & 0xFF)
	return color.NRGBA{R: r, G: gComp, B: b, A: 255}, nil
}

func (g *G3Image) resizeImage(w, h int) {
	if g.dc == nil {
		return
	}
	src := g.dc.Image()
	dst := image.NewRGBA(image.Rect(0, 0, w, h))
	var scaler xdraw.Scaler
	switch g.interpolationVal {
	case 1:
		scaler = xdraw.NearestNeighbor
	case 2:
		scaler = xdraw.BiLinear
	case 3:
		scaler = xdraw.CatmullRom
	default:
		scaler = xdraw.BiLinear
	}
	scaler.Scale(dst, dst.Bounds(), src, src.Bounds(), draw.Over, nil)
	g.dc = gg.NewContextForImage(dst)
}

func (g *G3Image) Crop(x0, y0, x1, y1 int) error {
	if g.dc == nil {
		return errors.New("no active context")
	}
	img := g.dc.Image()
	if x0 > x1 {
		x0, x1 = x1, x0
	}
	if y0 > y1 {
		y0, y1 = y1, y0
	}
	rect := image.Rect(x0, y0, x1, y1)
	rect = rect.Intersect(img.Bounds())
	if rect.Empty() {
		return errors.New("crop coordinates result in an empty image")
	}
	dst := image.NewRGBA(image.Rect(0, 0, rect.Dx(), rect.Dy()))
	draw.Draw(dst, dst.Bounds(), img, rect.Min, draw.Src)
	g.dc = gg.NewContextForImage(dst)
	return nil
}

func (g *G3Image) Sharpen(radius float64, amount float64) {
	if g.dc == nil {
		return
	}
	factor := amount
	if factor > 2.0 {
		factor = factor / 100.0
	}
	if factor <= 0 {
		return
	}
	src := g.dc.Image()
	bounds := src.Bounds()
	w, h := bounds.Dx(), bounds.Dy()
	dst := image.NewRGBA(image.Rect(0, 0, w, h))
	for y := range h {
		for x := range w {
			rSum, gSum, bSum := 0.0, 0.0, 0.0
			for _, dy := range []int{-1, 0, 1} {
				for _, dx := range []int{-1, 0, 1} {
					nx, ny := x+dx, y+dy
					if nx >= 0 && nx < w && ny >= 0 && ny < h {
						r, gVal, b, _ := src.At(nx, ny).RGBA()
						rf := float64(r >> 8)
						gf := float64(gVal >> 8)
						bf := float64(b >> 8)
						if dx == 0 && dy == 0 {
							rSum += rf * 5.0
							gSum += gf * 5.0
							bSum += bf * 5.0
						} else if dx == 0 || dy == 0 {
							rSum += rf * -1.0
							gSum += gf * -1.0
							bSum += bf * -1.0
						}
					}
				}
			}
			cr, cg, cb, ca := src.At(x, y).RGBA()
			crf := float64(cr >> 8)
			cgf := float64(cg >> 8)
			cbf := float64(cb >> 8)
			caf := uint8(ca >> 8)
			rf := crf + factor*(rSum-crf)
			gf := cgf + factor*(gSum-cgf)
			bf := cbf + factor*(bSum-cbf)
			ru := uint8(min(max(rf, 0), 255))
			gu := uint8(min(max(gf, 0), 255))
			bu := uint8(min(max(bf, 0), 255))
			dst.SetRGBA(x, y, color.RGBA{R: ru, G: gu, B: bu, A: caf})
		}
	}
	g.dc = gg.NewContextForImage(dst)
}

func (g *G3Image) getCanvas() Value {
	if g.canvasVal.Type == VTNativeObject {
		return g.canvasVal
	}
	canvas := &G3ImageCanvas{
		vm:    g.vm,
		image: g,
	}
	id := g.vm.nextDynamicNativeID
	g.vm.nextDynamicNativeID++
	canvas.objectID = id
	g.vm.g3imageCanvasItems[id] = canvas
	g.canvasVal = Value{Type: VTNativeObject, Num: id}
	return g.canvasVal
}

type G3ImageCanvas struct {
	vm       *VM
	objectID int64
	image    *G3Image
}

func (c *G3ImageCanvas) getFont() Value {
	if c.image.fontVal.Type == VTNativeObject {
		return c.image.fontVal
	}
	fontObj := &G3ImageFont{
		vm:     c.vm,
		image:  c.image,
		family: "Arial",
		size:   10,
		color:  color.Black,
	}
	id := c.vm.nextDynamicNativeID
	c.vm.nextDynamicNativeID++
	fontObj.objectID = id
	c.vm.g3imageFontItems[id] = fontObj
	c.image.fontVal = Value{Type: VTNativeObject, Num: id}
	return c.image.fontVal
}

func (c *G3ImageCanvas) getPen() Value {
	if c.image.penVal.Type == VTNativeObject {
		return c.image.penVal
	}
	penObj := &G3ImagePen{
		vm:    c.vm,
		image: c.image,
		color: color.Black,
		width: 1,
	}
	id := c.vm.nextDynamicNativeID
	c.vm.nextDynamicNativeID++
	penObj.objectID = id
	c.vm.g3imagePenItems[id] = penObj
	c.image.penVal = Value{Type: VTNativeObject, Num: id}
	return c.image.penVal
}

func (c *G3ImageCanvas) PrintText(x, y int, text string) error {
	if c.image.dc == nil {
		return errors.New("image context not initialized")
	}
	fontVal := c.getFont()
	fontObj := c.vm.g3imageFontItems[fontVal.Num]
	if err := fontObj.loadFont(); err != nil {
		return err
	}
	c.image.dc.SetColor(fontObj.color)
	c.image.dc.DrawString(text, float64(x), float64(y))
	return nil
}

func (c *G3ImageCanvas) DrawLine(x1, y1, x2, y2 int) error {
	if c.image.dc == nil {
		return errors.New("image context not initialized")
	}
	penVal := c.getPen()
	penObj := c.vm.g3imagePenItems[penVal.Num]
	c.image.dc.SetColor(penObj.color)
	c.image.dc.SetLineWidth(penObj.width)
	c.image.dc.DrawLine(float64(x1), float64(y1), float64(x2), float64(y2))
	c.image.dc.Stroke()
	return nil
}

func (c *G3ImageCanvas) DrawBar(x1, y1, x2, y2 int) error {
	if c.image.dc == nil {
		return errors.New("image context not initialized")
	}
	penVal := c.getPen()
	penObj := c.vm.g3imagePenItems[penVal.Num]
	c.image.dc.SetColor(penObj.color)
	w := float64(x2 - x1)
	h := float64(y2 - y1)
	c.image.dc.DrawRectangle(float64(x1), float64(y1), w, h)
	c.image.dc.Fill()
	return nil
}

type G3ImageFont struct {
	vm       *VM
	objectID int64
	image    *G3Image
	family   string
	size     float64
	color    color.Color
	bold     bool
	italic   bool
}

func (f *G3ImageFont) loadFont() error {
	fontPath := f.family
	lowerFamily := strings.ToLower(f.family)
	isBold := f.bold
	isItalic := f.italic
	var filename string
	switch lowerFamily {
	case "arial":
		if isBold && isItalic {
			filename = "arialbi.ttf"
		} else if isBold {
			filename = "arialbd.ttf"
		} else if isItalic {
			filename = "ariali.ttf"
		} else {
			filename = "arial.ttf"
		}
	case "times new roman", "times":
		if isBold && isItalic {
			filename = "timesbi.ttf"
		} else if isBold {
			filename = "timesbd.ttf"
		} else if isItalic {
			filename = "timesi.ttf"
		} else {
			filename = "times.ttf"
		}
	case "courier new", "courier":
		if isBold && isItalic {
			filename = "courbi.ttf"
		} else if isBold {
			filename = "courbd.ttf"
		} else if isItalic {
			filename = "couri.ttf"
		} else {
			filename = "cour.ttf"
		}
	case "verdana":
		if isBold && isItalic {
			filename = "verdanaz.ttf"
		} else if isBold {
			filename = "verdanab.ttf"
		} else if isItalic {
			filename = "verdanai.ttf"
		} else {
			filename = "verdana.ttf"
		}
	case "tahoma":
		if isBold {
			filename = "tahomabd.ttf"
		} else {
			filename = "tahoma.ttf"
		}
	}
	if filename != "" {
		winFonts := filepath.Join(os.Getenv("SystemRoot"), "Fonts", filename)
		if _, err := os.Stat(winFonts); err == nil {
			fontPath = winFonts
		} else {
			arialFallback := filepath.Join(os.Getenv("SystemRoot"), "Fonts", "arial.ttf")
			if _, err := os.Stat(arialFallback); err == nil {
				fontPath = arialFallback
			}
		}
	}
	if !filepath.IsAbs(fontPath) && !strings.Contains(fontPath, ":") {
		if resolved, err := f.image.resolveRootPath(fontPath); err == nil {
			fontPath = resolved
		}
	}
	face, err := gg.LoadFontFace(fontPath, f.size)
	if err != nil {
		return err
	}
	f.image.lastFontFace = face
	if f.image.dc != nil {
		f.image.dc.SetFontFace(face)
	}
	return nil
}

type G3ImagePen struct {
	vm       *VM
	objectID int64
	image    *G3Image
	color    color.Color
	width    float64
}

func (vm *VM) dispatchG3ImageCanvasMethod(c *G3ImageCanvas, member string, args []Value) Value {
	switch strings.ToLower(member) {
	case "printtext":
		if len(args) < 3 {
			vm.raise(vbscript.InternalError, NewAxonASPError(ErrG3ImageInvalidArgCount, nil, AxonASPErrorMessages[ErrG3ImageInvalidArgCount], "axonvm/lib_g3image.go", 0).Error())
			return NewEmpty()
		}
		x := int(vm.asInt(args[0]))
		y := int(vm.asInt(args[1]))
		text := args[2].String()
		err := c.PrintText(x, y, text)
		if err != nil {
			vm.raise(vbscript.InternalError, NewAxonASPError(ErrInvalidProcedureCallOrArg, err, err.Error(), "axonvm/lib_g3image.go", 0).Error())
		}
		return NewEmpty()
	case "drawline":
		if len(args) < 4 {
			vm.raise(vbscript.InternalError, NewAxonASPError(ErrG3ImageInvalidArgCount, nil, AxonASPErrorMessages[ErrG3ImageInvalidArgCount], "axonvm/lib_g3image.go", 0).Error())
			return NewEmpty()
		}
		x1 := int(vm.asInt(args[0]))
		y1 := int(vm.asInt(args[1]))
		x2 := int(vm.asInt(args[2]))
		y2 := int(vm.asInt(args[3]))
		err := c.DrawLine(x1, y1, x2, y2)
		if err != nil {
			vm.raise(vbscript.InternalError, NewAxonASPError(ErrInvalidProcedureCallOrArg, err, err.Error(), "axonvm/lib_g3image.go", 0).Error())
		}
		return NewEmpty()
	case "drawbar":
		if len(args) < 4 {
			vm.raise(vbscript.InternalError, NewAxonASPError(ErrG3ImageInvalidArgCount, nil, AxonASPErrorMessages[ErrG3ImageInvalidArgCount], "axonvm/lib_g3image.go", 0).Error())
			return NewEmpty()
		}
		x1 := int(vm.asInt(args[0]))
		y1 := int(vm.asInt(args[1]))
		x2 := int(vm.asInt(args[2]))
		y2 := int(vm.asInt(args[3]))
		err := c.DrawBar(x1, y1, x2, y2)
		if err != nil {
			vm.raise(vbscript.InternalError, NewAxonASPError(ErrInvalidProcedureCallOrArg, err, err.Error(), "axonvm/lib_g3image.go", 0).Error())
		}
		return NewEmpty()
	}
	return NewEmpty()
}

func (vm *VM) dispatchG3ImageCanvasPropertyGet(c *G3ImageCanvas, member string) Value {
	switch strings.ToLower(member) {
	case "font":
		return c.getFont()
	case "pen":
		return c.getPen()
	}
	return NewEmpty()
}

func (vm *VM) dispatchG3ImageFontPropertyGet(f *G3ImageFont, member string) Value {
	switch strings.ToLower(member) {
	case "family":
		return NewString(f.family)
	case "size":
		return NewDouble(f.size)
	case "bold":
		return NewBool(f.bold)
	case "italic":
		return NewBool(f.italic)
	case "color":
		r, gComp, b, _ := f.color.RGBA()
		ru := uint32(r >> 8)
		gu := uint32(gComp >> 8)
		bu := uint32(b >> 8)
		val := (bu << 16) | (gu << 8) | ru
		return NewInteger(int64(val))
	}
	return NewEmpty()
}

func (vm *VM) dispatchG3ImageFontPropertySet(f *G3ImageFont, member string, val Value) bool {
	switch strings.ToLower(member) {
	case "family":
		f.family = val.String()
		return true
	case "size":
		f.size = vm.asFloat(val)
		return true
	case "bold":
		f.bold = vm.asBool(val)
		return true
	case "italic":
		f.italic = vm.asBool(val)
		return true
	case "color":
		if c, err := f.image.parseColorVal(val); err == nil {
			f.color = c
		}
		return true
	}
	return false
}

func (vm *VM) dispatchG3ImagePenPropertyGet(p *G3ImagePen, member string) Value {
	switch strings.ToLower(member) {
	case "color":
		r, gComp, b, _ := p.color.RGBA()
		ru := uint32(r >> 8)
		gu := uint32(gComp >> 8)
		bu := uint32(b >> 8)
		val := (bu << 16) | (gu << 8) | ru
		return NewInteger(int64(val))
	case "width":
		return NewDouble(p.width)
	}
	return NewEmpty()
}

func (vm *VM) dispatchG3ImagePenPropertySet(p *G3ImagePen, member string, val Value) bool {
	switch strings.ToLower(member) {
	case "color":
		if c, err := p.image.parseColorVal(val); err == nil {
			p.color = c
		}
		return true
	case "width":
		p.width = vm.asFloat(val)
		return true
	}
	return false
}
