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
 * This library integrates github.com/fogleman/gg.
 * gg is licensed under the MIT License.
 * Copyright (c) 2016 Michael Fogleman.
 */
package server

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
	"reflect"
	"strings"

	"github.com/fogleman/gg"
	"golang.org/x/image/font"
)

type G3IMAGE struct {
	ctx           *ExecutionContext
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

type GGImageValue struct {
	img image.Image
}

type GGPatternValue struct {
	pattern  gg.Pattern
	gradient gg.Gradient
}

type GGMaskValue struct {
	mask *image.Alpha
}

type GGMatrixValue struct {
	matrix gg.Matrix
}

type GGPointValue struct {
	point gg.Point
}

type GGFontFaceValue struct {
	face font.Face
}

func (g *G3IMAGE) ensureDefaults() {
	if g.defaultFormat == "" {
		g.defaultFormat = "png"
	}
	if g.jpgQuality <= 0 || g.jpgQuality > 100 {
		g.jpgQuality = 90
	}
}

func (g *G3IMAGE) GetProperty(name string) interface{} {
	g.ensureDefaults()
	switch strings.ToLower(name) {
	case "hascontext":
		return g.dc != nil
	case "width":
		if g.dc == nil {
			return 0
		}
		return g.dc.Width()
	case "height":
		if g.dc == nil {
			return 0
		}
		return g.dc.Height()
	case "lasterror":
		return g.lastErr
	case "lastmimetype", "contenttype", "mimetype":
		return g.lastMimeType
	case "lasttempfile", "tempfile":
		return g.lastTempFile
	case "lastbytes", "content":
		if g.lastBytes == nil {
			return []byte{}
		}
		buf := make([]byte, len(g.lastBytes))
		copy(buf, g.lastBytes)
		return buf
	case "defaultformat":
		return g.defaultFormat
	case "jpgquality", "jpegquality":
		return g.jpgQuality
	case "alignleft":
		return int(gg.AlignLeft)
	case "aligncenter":
		return int(gg.AlignCenter)
	case "alignright":
		return int(gg.AlignRight)
	case "fillrulewinding":
		return int(gg.FillRuleWinding)
	case "fillruleevenodd":
		return int(gg.FillRuleEvenOdd)
	case "linecapround":
		return int(gg.LineCapRound)
	case "linecapbutt":
		return int(gg.LineCapButt)
	case "linecapsquare":
		return int(gg.LineCapSquare)
	case "linejoinround":
		return int(gg.LineJoinRound)
	case "linejoinbevel":
		return int(gg.LineJoinBevel)
	case "repeatboth":
		return int(gg.RepeatBoth)
	case "repeatx":
		return int(gg.RepeatX)
	case "repeaty":
		return int(gg.RepeatY)
	case "repeatnone":
		return int(gg.RepeatNone)
	default:
		return nil
	}
}

func (g *G3IMAGE) SetProperty(name string, value interface{}) {
	g.ensureDefaults()
	switch strings.ToLower(name) {
	case "defaultformat":
		format := strings.ToLower(strings.TrimSpace(fmt.Sprintf("%v", value)))
		if format == "png" || format == "jpg" || format == "jpeg" {
			g.defaultFormat = format
		}
	case "jpgquality", "jpegquality":
		q := toInt(value)
		if q < 1 {
			q = 1
		}
		if q > 100 {
			q = 100
		}
		g.jpgQuality = q
	}
}

func (g *G3IMAGE) CallMethod(name string, args ...interface{}) interface{} {
	g.ensureDefaults()
	method := strings.ToLower(strings.TrimSpace(name))

	switch method {
	case "close", "dispose", "release", "destroy", "reset", "clearcontext", "clearimage":
		g.releaseResources(true)
		g.clearError()
		return true

	case "new", "newcontext", "create", "createcontext", "init":
		if len(args) < 2 {
			g.setError("newcontext requires width and height")
			return nil
		}
		g.releaseResources(false)
		w := toInt(args[0])
		h := toInt(args[1])
		if w <= 0 || h <= 0 {
			g.setError("newcontext requires positive dimensions")
			return nil
		}
		g.dc = gg.NewContext(w, h)
		g.clearError()
		return g

	case "newcontextforimage", "contextforimage", "useimage", "setimage":
		if len(args) < 1 {
			g.setError("newcontextforimage requires an image or path")
			return nil
		}
		g.releaseResources(false)
		im, err := g.asImage(args[0])
		if err != nil {
			g.setError(err.Error())
			return nil
		}
		g.dc = gg.NewContextForImage(im)
		g.lastLoaded = im
		g.clearError()
		return g

	case "newcontextforrgba", "contextforrgba":
		if len(args) < 1 {
			g.setError("newcontextforrgba requires an image or path")
			return nil
		}
		g.releaseResources(false)
		im, err := g.asImage(args[0])
		if err != nil {
			g.setError(err.Error())
			return nil
		}
		rgba := image.NewRGBA(im.Bounds())
		draw.Draw(rgba, rgba.Bounds(), im, im.Bounds().Min, draw.Src)
		g.dc = gg.NewContextForRGBA(rgba)
		g.lastLoaded = im
		g.clearError()
		return g

	case "loadimage", "load":
		if len(args) < 1 {
			g.setError("loadimage requires path")
			return nil
		}
		im, err := g.loadImageFromPath(fmt.Sprintf("%v", args[0]), "")
		if err != nil {
			g.setError(err.Error())
			return nil
		}
		g.lastLoaded = im
		g.clearError()
		return &GGImageValue{img: im}

	case "loadpng":
		if len(args) < 1 {
			g.setError("loadpng requires path")
			return nil
		}
		im, err := g.loadImageFromPath(fmt.Sprintf("%v", args[0]), "png")
		if err != nil {
			g.setError(err.Error())
			return nil
		}
		g.lastLoaded = im
		g.clearError()
		return &GGImageValue{img: im}

	case "loadjpg", "loadjpeg":
		if len(args) < 1 {
			g.setError("loadjpg requires path")
			return nil
		}
		im, err := g.loadImageFromPath(fmt.Sprintf("%v", args[0]), "jpg")
		if err != nil {
			g.setError(err.Error())
			return nil
		}
		g.lastLoaded = im
		g.clearError()
		return &GGImageValue{img: im}

	case "image", "getimage":
		if g.dc == nil {
			g.setError("no active context")
			return nil
		}
		g.clearError()
		return &GGImageValue{img: g.dc.Image()}

	case "newlineargradient":
		if len(args) < 4 {
			g.setError("newlineargradient requires x0, y0, x1, y1")
			return nil
		}
		grad := gg.NewLinearGradient(toFloat(args[0]), toFloat(args[1]), toFloat(args[2]), toFloat(args[3]))
		g.clearError()
		return &GGPatternValue{pattern: grad, gradient: grad}

	case "newradialgradient":
		if len(args) < 6 {
			g.setError("newradialgradient requires x0, y0, r0, x1, y1, r1")
			return nil
		}
		grad := gg.NewRadialGradient(toFloat(args[0]), toFloat(args[1]), toFloat(args[2]), toFloat(args[3]), toFloat(args[4]), toFloat(args[5]))
		g.clearError()
		return &GGPatternValue{pattern: grad, gradient: grad}

	case "newsolidpattern":
		if len(args) < 1 {
			g.setError("newsolidpattern requires color")
			return nil
		}
		c, err := parseColorArg(args[0])
		if err != nil {
			g.setError(err.Error())
			return nil
		}
		p := gg.NewSolidPattern(c)
		g.clearError()
		return &GGPatternValue{pattern: p}

	case "newsurfacepattern":
		if len(args) < 1 {
			g.setError("newsurfacepattern requires image")
			return nil
		}
		im, err := g.asImage(args[0])
		if err != nil {
			g.setError(err.Error())
			return nil
		}
		op := gg.RepeatBoth
		if len(args) > 1 {
			op = gg.RepeatOp(toInt(args[1]))
		}
		p := gg.NewSurfacePattern(im, op)
		g.clearError()
		return &GGPatternValue{pattern: p}

	case "identitymatrix":
		return &GGMatrixValue{matrix: gg.Identity()}
	case "translate":
		if len(args) == 2 && g.dc == nil {
			return &GGMatrixValue{matrix: gg.Translate(toFloat(args[0]), toFloat(args[1]))}
		}
	case "scale":
		if len(args) == 2 && g.dc == nil {
			return &GGMatrixValue{matrix: gg.Scale(toFloat(args[0]), toFloat(args[1]))}
		}
	case "rotate":
		if len(args) == 1 && g.dc == nil {
			return &GGMatrixValue{matrix: gg.Rotate(toFloat(args[0]))}
		}
	case "shear":
		if len(args) == 2 && g.dc == nil {
			return &GGMatrixValue{matrix: gg.Shear(toFloat(args[0]), toFloat(args[1]))}
		}

	case "quadraticbezier":
		if len(args) < 6 {
			g.setError("quadraticbezier requires 6 arguments")
			return nil
		}
		points := gg.QuadraticBezier(toFloat(args[0]), toFloat(args[1]), toFloat(args[2]), toFloat(args[3]), toFloat(args[4]), toFloat(args[5]))
		return wrapPoints(points)

	case "cubicbezier":
		if len(args) < 8 {
			g.setError("cubicbezier requires 8 arguments")
			return nil
		}
		points := gg.CubicBezier(toFloat(args[0]), toFloat(args[1]), toFloat(args[2]), toFloat(args[3]), toFloat(args[4]), toFloat(args[5]), toFloat(args[6]), toFloat(args[7]))
		return wrapPoints(points)

	case "radians":
		if len(args) < 1 {
			return float64(0)
		}
		return gg.Radians(toFloat(args[0]))
	case "degrees":
		if len(args) < 1 {
			return float64(0)
		}
		return gg.Degrees(toFloat(args[0]))

	case "loadfontface":
		if len(args) < 2 {
			g.setError("loadfontface requires path and points")
			return nil
		}
		fontPath, err := g.resolveRootPath(fmt.Sprintf("%v", args[0]))
		if err != nil {
			g.setError(err.Error())
			return nil
		}
		f, err := gg.LoadFontFace(fontPath, toFloat(args[1]))
		if err != nil {
			g.setError(err.Error())
			return nil
		}
		g.lastFontFace = f
		if g.dc != nil {
			g.dc.SetFontFace(f)
		}
		g.clearError()
		return &GGFontFaceValue{face: f}

	case "savepng":
		if g.dc == nil {
			g.setError("no active context")
			return false
		}
		if len(args) < 1 {
			g.setError("savepng requires path")
			return false
		}
		err := g.savePNGToPath(fmt.Sprintf("%v", args[0]))
		if err != nil {
			g.setError(err.Error())
			return false
		}
		g.clearError()
		return true

	case "savejpg", "savejpeg":
		if g.dc == nil {
			g.setError("no active context")
			return false
		}
		if len(args) < 1 {
			g.setError("savejpg requires path")
			return false
		}
		quality := g.jpgQuality
		if len(args) > 1 {
			quality = toInt(args[1])
		}
		err := g.saveJPGToPath(fmt.Sprintf("%v", args[0]), quality)
		if err != nil {
			g.setError(err.Error())
			return false
		}
		g.clearError()
		return true

	case "encodepng", "topngbytes", "renderpngbytes", "renderpng":
		data, err := g.renderPNGBytes()
		if err != nil {
			g.setError(err.Error())
			return nil
		}
		g.clearError()
		return data

	case "encodejpg", "tojpgbytes", "tojpegbytes", "renderjpgbytes", "renderjpegbytes":
		quality := g.jpgQuality
		if len(args) > 0 {
			quality = toInt(args[0])
		}
		data, err := g.renderJPGBytes(quality)
		if err != nil {
			g.setError(err.Error())
			return nil
		}
		g.clearError()
		return data

	case "rendercontent", "getcontent", "contentforresponse", "rendertobrowser":
		format := g.defaultFormat
		quality := g.jpgQuality
		if len(args) > 0 {
			format = strings.ToLower(strings.TrimSpace(fmt.Sprintf("%v", args[0])))
		}
		if len(args) > 1 {
			quality = toInt(args[1])
		}
		data, err := g.renderContent(format, quality)
		if err != nil {
			g.setError(err.Error())
			return nil
		}
		g.clearError()
		return data

	case "renderviatemp", "renderbytemp", "getcontentviatemp", "rendertemp":
		format := g.defaultFormat
		quality := g.jpgQuality
		if len(args) > 0 {
			format = strings.ToLower(strings.TrimSpace(fmt.Sprintf("%v", args[0])))
		}
		if len(args) > 1 {
			quality = toInt(args[1])
		}
		data, err := g.renderViaTemp(format, quality)
		if err != nil {
			g.setError(err.Error())
			return nil
		}
		g.clearError()
		return data

	case "lastmimetype", "getlastmimetype":
		return g.lastMimeType
	case "lasterror", "getlasterror":
		return g.lastErr
	}

	if g.dc == nil {
		g.setError("no active context")
		return nil
	}

	result, err := g.invokeContextMethod(method, args)
	if err != nil {
		g.setError(err.Error())
		return nil
	}
	g.clearError()
	return result
}

func (g *G3IMAGE) releaseResources(clearError bool) {
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

func (g *G3IMAGE) invokeContextMethod(method string, args []interface{}) (interface{}, error) {
	ctxMethod, ok := ggContextMethodMap[method]
	if !ok {
		ctxMethod = canonicalMethodName(method)
	}

	receiver := reflect.ValueOf(g.dc)
	m := receiver.MethodByName(ctxMethod)
	if !m.IsValid() {
		return nil, fmt.Errorf("unsupported method: %s", method)
	}

	mt := m.Type()
	converted := make([]reflect.Value, 0, len(args))

	if mt.IsVariadic() {
		fixedCount := mt.NumIn() - 1
		if len(args) < fixedCount {
			return nil, fmt.Errorf("%s requires at least %d arguments", method, fixedCount)
		}
		for i := 0; i < fixedCount; i++ {
			v, err := g.convertContextArg(args[i], mt.In(i))
			if err != nil {
				return nil, err
			}
			converted = append(converted, v)
		}
		variadicType := mt.In(mt.NumIn() - 1).Elem()
		for i := fixedCount; i < len(args); i++ {
			v, err := g.convertContextArg(args[i], variadicType)
			if err != nil {
				return nil, err
			}
			converted = append(converted, v)
		}
	} else {
		if len(args) != mt.NumIn() {
			return nil, fmt.Errorf("%s requires %d arguments", method, mt.NumIn())
		}
		for i := 0; i < mt.NumIn(); i++ {
			v, err := g.convertContextArg(args[i], mt.In(i))
			if err != nil {
				return nil, err
			}
			converted = append(converted, v)
		}
	}

	outs := m.Call(converted)
	return g.normalizeReturnValues(outs)
}

func (g *G3IMAGE) convertContextArg(input interface{}, target reflect.Type) (reflect.Value, error) {
	if target == reflect.TypeOf((*image.Image)(nil)).Elem() {
		im, err := g.asImage(input)
		if err != nil {
			return reflect.Value{}, err
		}
		return reflect.ValueOf(im), nil
	}

	if target == reflect.TypeOf((*gg.Pattern)(nil)).Elem() {
		if p, ok := input.(*GGPatternValue); ok && p != nil && p.pattern != nil {
			return reflect.ValueOf(p.pattern), nil
		}
		return reflect.Value{}, errors.New("expected pattern object")
	}

	if target == reflect.TypeOf((*image.Alpha)(nil)) {
		if mask, ok := input.(*GGMaskValue); ok && mask != nil && mask.mask != nil {
			return reflect.ValueOf(mask.mask), nil
		}
		return reflect.Value{}, errors.New("expected mask object")
	}

	if target == reflect.TypeOf((*font.Face)(nil)).Elem() {
		if fv, ok := input.(*GGFontFaceValue); ok && fv != nil && fv.face != nil {
			return reflect.ValueOf(fv.face), nil
		}
		if g.lastFontFace != nil {
			return reflect.ValueOf(g.lastFontFace), nil
		}
		return reflect.Value{}, errors.New("expected font face object")
	}

	if target == reflect.TypeOf((*color.Color)(nil)).Elem() {
		c, err := parseColorArg(input)
		if err != nil {
			return reflect.Value{}, err
		}
		return reflect.ValueOf(c), nil
	}

	if target.PkgPath() == "github.com/fogleman/gg" {
		switch target.Name() {
		case "Align":
			return reflect.ValueOf(gg.Align(toInt(input))), nil
		case "LineCap":
			return reflect.ValueOf(gg.LineCap(toInt(input))), nil
		case "LineJoin":
			return reflect.ValueOf(gg.LineJoin(toInt(input))), nil
		case "FillRule":
			return reflect.ValueOf(gg.FillRule(toInt(input))), nil
		}
	}

	switch target.Kind() {
	case reflect.String:
		return reflect.ValueOf(fmt.Sprintf("%v", input)).Convert(target), nil
	case reflect.Int, reflect.Int8, reflect.Int16, reflect.Int32, reflect.Int64:
		v := reflect.New(target).Elem()
		v.SetInt(int64(toInt(input)))
		return v, nil
	case reflect.Float32, reflect.Float64:
		v := reflect.New(target).Elem()
		v.SetFloat(toFloat(input))
		return v, nil
	case reflect.Bool:
		v := reflect.New(target).Elem()
		v.SetBool(toBool(input))
		return v, nil
	}

	iv := reflect.ValueOf(input)
	if iv.IsValid() {
		if iv.Type().AssignableTo(target) {
			return iv, nil
		}
		if iv.Type().ConvertibleTo(target) {
			return iv.Convert(target), nil
		}
	}

	return reflect.Value{}, fmt.Errorf("cannot convert %T to %s", input, target.String())
}

func (g *G3IMAGE) normalizeReturnValues(outs []reflect.Value) (interface{}, error) {
	if len(outs) == 0 {
		return true, nil
	}

	if len(outs) == 1 {
		if err, ok := outs[0].Interface().(error); ok {
			if err != nil {
				return nil, err
			}
			return true, nil
		}
		return normalizeGGReturn(outs[0].Interface()), nil
	}

	if lastErr, ok := outs[len(outs)-1].Interface().(error); ok {
		if lastErr != nil {
			return nil, lastErr
		}
		outs = outs[:len(outs)-1]
	}

	result := make([]interface{}, 0, len(outs))
	for _, out := range outs {
		result = append(result, normalizeGGReturn(out.Interface()))
	}
	return result, nil
}

func normalizeGGReturn(v interface{}) interface{} {
	switch t := v.(type) {
	case nil:
		return nil
	case image.Image:
		return &GGImageValue{img: t}
	case *image.Alpha:
		return &GGMaskValue{mask: t}
	case gg.Pattern:
		return &GGPatternValue{pattern: t}
	case gg.Gradient:
		return &GGPatternValue{pattern: t, gradient: t}
	case gg.Matrix:
		return &GGMatrixValue{matrix: t}
	case gg.Point:
		return &GGPointValue{point: t}
	case []gg.Point:
		return wrapPoints(t)
	case []string:
		list := make([]interface{}, 0, len(t))
		for _, item := range t {
			list = append(list, item)
		}
		return list
	default:
		return v
	}
}

func (g *G3IMAGE) renderPNGBytes() ([]byte, error) {
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

func (g *G3IMAGE) renderJPGBytes(quality int) ([]byte, error) {
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

func (g *G3IMAGE) renderContent(format string, quality int) ([]byte, error) {
	if g.dc == nil {
		return nil, errors.New("no active context")
	}
	format = normalizeImageFormat(format)
	if format == "jpg" {
		data, err := g.renderJPGBytes(quality)
		if err == nil {
			return data, nil
		}
		return g.renderViaTemp("jpg", quality)
	}
	data, err := g.renderPNGBytes()
	if err == nil {
		return data, nil
	}
	return g.renderViaTemp("png", quality)
}

func (g *G3IMAGE) renderViaTemp(format string, quality int) ([]byte, error) {
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

func (g *G3IMAGE) savePNGToPath(path string) error {
	fullPath, err := g.resolveRootPath(path)
	if err != nil {
		return err
	}
	return g.savePNGAbsolute(fullPath)
}

func (g *G3IMAGE) saveJPGToPath(path string, quality int) error {
	fullPath, err := g.resolveRootPath(path)
	if err != nil {
		return err
	}
	return g.saveJPGAbsolute(fullPath, quality)
}

func (g *G3IMAGE) savePNGAbsolute(path string) error {
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

func (g *G3IMAGE) saveJPGAbsolute(path string, quality int) error {
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

func (g *G3IMAGE) resolveRootPath(rel string) (string, error) {
	rel = strings.TrimSpace(rel)
	if rel == "" {
		return "", errors.New("path is required")
	}

	if g.ctx == nil {
		return filepath.Abs(rel)
	}

	mapped := g.ctx.Server_MapPath(rel)
	if mapped == "" {
		return "", errors.New("invalid mapped path")
	}

	rootDir, _ := filepath.Abs(g.ctx.RootDir)
	absPath, _ := filepath.Abs(mapped)

	if !strings.HasPrefix(strings.ToLower(absPath), strings.ToLower(rootDir)) {
		return "", errors.New("access denied outside root directory")
	}

	return absPath, nil
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

func (g *G3IMAGE) loadImageFromPath(path string, kind string) (image.Image, error) {
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

func (g *G3IMAGE) asImage(value interface{}) (image.Image, error) {
	switch v := value.(type) {
	case image.Image:
		return v, nil
	case *GGImageValue:
		if v != nil && v.img != nil {
			return v.img, nil
		}
	case *G3IMAGE:
		if v != nil && v.dc != nil {
			return v.dc.Image(), nil
		}
	case string:
		return g.loadImageFromPath(v, "")
	}
	return nil, fmt.Errorf("invalid image source: %T", value)
}

func (g *G3IMAGE) clearError() {
	g.lastErr = ""
}

func (g *G3IMAGE) setError(err string) {
	g.lastErr = err
}

func wrapPoints(points []gg.Point) []interface{} {
	result := make([]interface{}, 0, len(points))
	for _, p := range points {
		cp := p
		result = append(result, &GGPointValue{point: cp})
	}
	return result
}

func canonicalMethodName(method string) string {
	parts := strings.FieldsFunc(strings.ToLower(method), func(r rune) bool {
		return r == '_' || r == '-' || r == ' '
	})
	if len(parts) == 0 {
		return ""
	}
	for i := range parts {
		if len(parts[i]) == 0 {
			continue
		}
		parts[i] = strings.ToUpper(parts[i][:1]) + parts[i][1:]
	}
	return strings.Join(parts, "")
}

var ggContextMethodMap = map[string]string{
	"asmask":                 "AsMask",
	"clear":                  "Clear",
	"clearpath":              "ClearPath",
	"clip":                   "Clip",
	"clippreserve":           "ClipPreserve",
	"closepath":              "ClosePath",
	"cubicto":                "CubicTo",
	"drawarc":                "DrawArc",
	"drawcircle":             "DrawCircle",
	"drawellipse":            "DrawEllipse",
	"drawellipticalarc":      "DrawEllipticalArc",
	"drawimage":              "DrawImage",
	"drawimageanchored":      "DrawImageAnchored",
	"drawline":               "DrawLine",
	"drawpoint":              "DrawPoint",
	"drawrectangle":          "DrawRectangle",
	"drawregularpolygon":     "DrawRegularPolygon",
	"drawroundedrectangle":   "DrawRoundedRectangle",
	"drawstring":             "DrawString",
	"drawstringanchored":     "DrawStringAnchored",
	"drawstringwrapped":      "DrawStringWrapped",
	"encodepng":              "EncodePNG",
	"fill":                   "Fill",
	"fillpreserve":           "FillPreserve",
	"fontheight":             "FontHeight",
	"getcurrentpoint":        "GetCurrentPoint",
	"height":                 "Height",
	"identity":               "Identity",
	"image":                  "Image",
	"invertmask":             "InvertMask",
	"inverty":                "InvertY",
	"lineto":                 "LineTo",
	"loadfontface":           "LoadFontFace",
	"measuremultilinestring": "MeasureMultilineString",
	"measurestring":          "MeasureString",
	"moveto":                 "MoveTo",
	"newsubpath":             "NewSubPath",
	"pop":                    "Pop",
	"push":                   "Push",
	"quadraticto":            "QuadraticTo",
	"resetclip":              "ResetClip",
	"rotate":                 "Rotate",
	"rotateabout":            "RotateAbout",
	"savepng":                "SavePNG",
	"scale":                  "Scale",
	"scaleabout":             "ScaleAbout",
	"setcolor":               "SetColor",
	"setdash":                "SetDash",
	"setdashoffset":          "SetDashOffset",
	"setfillrule":            "SetFillRule",
	"setfillruleevenodd":     "SetFillRuleEvenOdd",
	"setfillrulewinding":     "SetFillRuleWinding",
	"setfillstyle":           "SetFillStyle",
	"setfontface":            "SetFontFace",
	"sethexcolor":            "SetHexColor",
	"setlinecap":             "SetLineCap",
	"setlinecapbutt":         "SetLineCapButt",
	"setlinecapround":        "SetLineCapRound",
	"setlinecapsquare":       "SetLineCapSquare",
	"setlinejoin":            "SetLineJoin",
	"setlinejoinbevel":       "SetLineJoinBevel",
	"setlinejoinround":       "SetLineJoinRound",
	"setlinewidth":           "SetLineWidth",
	"setmask":                "SetMask",
	"setpixel":               "SetPixel",
	"setrgb":                 "SetRGB",
	"setrgb255":              "SetRGB255",
	"setrgba":                "SetRGBA",
	"setrgba255":             "SetRGBA255",
	"setstrokestyle":         "SetStrokeStyle",
	"shear":                  "Shear",
	"shearabout":             "ShearAbout",
	"stroke":                 "Stroke",
	"strokepreserve":         "StrokePreserve",
	"transformpoint":         "TransformPoint",
	"translate":              "Translate",
	"width":                  "Width",
	"wordwrap":               "WordWrap",
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

func parseColorArg(value interface{}) (color.Color, error) {
	switch v := value.(type) {
	case color.Color:
		return v, nil
	case string:
		return parseColorString(v)
	case []interface{}:
		if len(v) == 3 || len(v) == 4 {
			r := uint8(toInt(v[0]))
			g := uint8(toInt(v[1]))
			b := uint8(toInt(v[2]))
			a := uint8(255)
			if len(v) == 4 {
				a = uint8(toInt(v[3]))
			}
			return color.NRGBA{R: r, G: g, B: b, A: a}, nil
		}
	}
	return nil, fmt.Errorf("unsupported color value: %T", value)
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
		r := uint8(toInt(strings.TrimSpace(parts[0])))
		g := uint8(toInt(strings.TrimSpace(parts[1])))
		b := uint8(toInt(strings.TrimSpace(parts[2])))
		a := uint8(255)
		if len(parts) == 4 {
			a = uint8(toInt(strings.TrimSpace(parts[3])))
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

func (v *GGImageValue) GetProperty(name string) interface{} {
	if v == nil || v.img == nil {
		return nil
	}
	switch strings.ToLower(name) {
	case "width":
		return v.img.Bounds().Dx()
	case "height":
		return v.img.Bounds().Dy()
	case "bounds":
		b := v.img.Bounds()
		return []interface{}{b.Min.X, b.Min.Y, b.Max.X, b.Max.Y}
	}
	return nil
}

func (v *GGImageValue) SetProperty(name string, value interface{}) error {
	return nil
}

func (v *GGImageValue) CallMethod(name string, args ...interface{}) (interface{}, error) {
	switch strings.ToLower(name) {
	case "width":
		return v.GetProperty("width"), nil
	case "height":
		return v.GetProperty("height"), nil
	}
	return nil, nil
}

func (p *GGPatternValue) GetProperty(name string) interface{} {
	if p == nil {
		return nil
	}
	switch strings.ToLower(name) {
	case "isgradient":
		return p.gradient != nil
	}
	return nil
}

func (p *GGPatternValue) SetProperty(name string, value interface{}) error {
	return nil
}

func (p *GGPatternValue) CallMethod(name string, args ...interface{}) (interface{}, error) {
	switch strings.ToLower(name) {
	case "addcolorstop":
		if p.gradient == nil {
			return nil, errors.New("pattern is not a gradient")
		}
		if len(args) < 2 {
			return nil, errors.New("addcolorstop requires offset and color")
		}
		c, err := parseColorArg(args[1])
		if err != nil {
			return nil, err
		}
		p.gradient.AddColorStop(toFloat(args[0]), c)
		return p, nil
	}
	return nil, nil
}

func (m *GGMaskValue) GetProperty(name string) interface{} {
	if m == nil || m.mask == nil {
		return nil
	}
	switch strings.ToLower(name) {
	case "width":
		return m.mask.Bounds().Dx()
	case "height":
		return m.mask.Bounds().Dy()
	}
	return nil
}

func (m *GGMaskValue) SetProperty(name string, value interface{}) error {
	return nil
}

func (m *GGMaskValue) CallMethod(name string, args ...interface{}) (interface{}, error) {
	return nil, nil
}

func (m *GGMatrixValue) GetProperty(name string) interface{} {
	if m == nil {
		return nil
	}
	switch strings.ToLower(name) {
	case "xx":
		return m.matrix.XX
	case "yx":
		return m.matrix.YX
	case "xy":
		return m.matrix.XY
	case "yy":
		return m.matrix.YY
	case "x0":
		return m.matrix.X0
	case "y0":
		return m.matrix.Y0
	}
	return nil
}

func (m *GGMatrixValue) SetProperty(name string, value interface{}) error {
	switch strings.ToLower(name) {
	case "xx":
		m.matrix.XX = toFloat(value)
	case "yx":
		m.matrix.YX = toFloat(value)
	case "xy":
		m.matrix.XY = toFloat(value)
	case "yy":
		m.matrix.YY = toFloat(value)
	case "x0":
		m.matrix.X0 = toFloat(value)
	case "y0":
		m.matrix.Y0 = toFloat(value)
	}
	return nil
}

func (m *GGMatrixValue) CallMethod(name string, args ...interface{}) (interface{}, error) {
	if m == nil {
		return nil, errors.New("matrix is nil")
	}
	switch strings.ToLower(name) {
	case "multiply":
		if len(args) < 1 {
			return nil, errors.New("multiply requires matrix argument")
		}
		other, ok := args[0].(*GGMatrixValue)
		if !ok || other == nil {
			return nil, errors.New("invalid matrix argument")
		}
		m.matrix = m.matrix.Multiply(other.matrix)
		return m, nil
	case "translate":
		if len(args) < 2 {
			return nil, errors.New("translate requires x and y")
		}
		m.matrix = m.matrix.Translate(toFloat(args[0]), toFloat(args[1]))
		return m, nil
	case "scale":
		if len(args) < 2 {
			return nil, errors.New("scale requires x and y")
		}
		m.matrix = m.matrix.Scale(toFloat(args[0]), toFloat(args[1]))
		return m, nil
	case "rotate":
		if len(args) < 1 {
			return nil, errors.New("rotate requires angle")
		}
		m.matrix = m.matrix.Rotate(toFloat(args[0]))
		return m, nil
	case "shear":
		if len(args) < 2 {
			return nil, errors.New("shear requires x and y")
		}
		m.matrix = m.matrix.Shear(toFloat(args[0]), toFloat(args[1]))
		return m, nil
	case "transformpoint":
		if len(args) < 2 {
			return nil, errors.New("transformpoint requires x and y")
		}
		x, y := m.matrix.TransformPoint(toFloat(args[0]), toFloat(args[1]))
		return []interface{}{x, y}, nil
	case "transformvector":
		if len(args) < 2 {
			return nil, errors.New("transformvector requires x and y")
		}
		x, y := m.matrix.TransformVector(toFloat(args[0]), toFloat(args[1]))
		return []interface{}{x, y}, nil
	}
	return nil, nil
}

func (p *GGPointValue) GetProperty(name string) interface{} {
	if p == nil {
		return nil
	}
	switch strings.ToLower(name) {
	case "x":
		return p.point.X
	case "y":
		return p.point.Y
	}
	return nil
}

func (p *GGPointValue) SetProperty(name string, value interface{}) error {
	switch strings.ToLower(name) {
	case "x":
		p.point.X = toFloat(value)
	case "y":
		p.point.Y = toFloat(value)
	}
	return nil
}

func (p *GGPointValue) CallMethod(name string, args ...interface{}) (interface{}, error) {
	if p == nil {
		return nil, errors.New("point is nil")
	}
	switch strings.ToLower(name) {
	case "distance":
		if len(args) < 1 {
			return nil, errors.New("distance requires point")
		}
		other, ok := args[0].(*GGPointValue)
		if !ok || other == nil {
			return nil, errors.New("invalid point")
		}
		return p.point.Distance(other.point), nil
	case "interpolate":
		if len(args) < 2 {
			return nil, errors.New("interpolate requires point and t")
		}
		other, ok := args[0].(*GGPointValue)
		if !ok || other == nil {
			return nil, errors.New("invalid point")
		}
		return &GGPointValue{point: p.point.Interpolate(other.point, toFloat(args[1]))}, nil
	}
	return nil, nil
}

func (f *GGFontFaceValue) GetProperty(name string) interface{} {
	if f == nil || f.face == nil {
		return nil
	}
	return nil
}

func (f *GGFontFaceValue) SetProperty(name string, value interface{}) error {
	return nil
}

func (f *GGFontFaceValue) CallMethod(name string, args ...interface{}) (interface{}, error) {
	return nil, nil
}
