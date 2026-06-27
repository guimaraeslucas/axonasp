//go:build !wasm

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
	"testing"
)

func TestG3Image(t *testing.T) {
	vm := NewVM(nil, nil, 0)

	imgLib := vm.newG3ImageObject()
	if imgLib.Type != VTNativeObject {
		t.Fatalf("expected VTNativeObject, got %v", imgLib.Type)
	}

	obj := vm.g3imageItems[imgLib.Num]
	if obj == nil {
		t.Fatal("expected G3Image object")
	}

	// Test new context
	obj.DispatchMethod("New", []Value{NewInteger(100), NewInteger(100)})

	hasCtx := obj.DispatchPropertyGet("HasContext")
	if hasCtx.Type != VTBool || hasCtx.Num == 0 {
		t.Error("expected HasContext true")
	}

	width := obj.DispatchPropertyGet("Width")
	if width.Num != 100 {
		t.Errorf("expected width 100, got %d", width.Num)
	}

	// Test cleanup
	obj.DispatchMethod("Close", nil)
	hasCtx = obj.DispatchPropertyGet("HasContext")
	if hasCtx.Type != VTBool || hasCtx.Num != 0 {
		t.Error("expected HasContext false")
	}
	if _, exists := vm.g3imageItems[imgLib.Num]; exists {
		t.Fatal("expected image object to be detached from VM map after Close")
	}
	if obj.dc != nil || obj.lastLoaded != nil || obj.lastBytes != nil || obj.lastFontFace != nil {
		t.Fatal("expected Close to clear all image resource pointers")
	}
}

func TestPersitsJpegCompatibility(t *testing.T) {
	vm := NewVM(nil, nil, 0)

	// Test creation via progid (mocked in tests via direct call or VM dispatch)
	imgLib := vm.newPersitsJpegObject()
	obj := vm.g3imageItems[imgLib.Num]
	if obj == nil {
		t.Fatal("expected G3Image object")
	}
	if !obj.isPersits {
		t.Error("expected isPersits to be true")
	}

	// Test New(w, h, color) with White
	obj.DispatchMethod("New", []Value{NewInteger(200), NewInteger(150), NewString("#FFFFFF")})

	if obj.originalWidth != 200 || obj.originalHeight != 150 {
		t.Errorf("expected original dimensions 200x150, got %dx%d", obj.originalWidth, obj.originalHeight)
	}

	// Test Quality get/set
	obj.DispatchPropertySet("Quality", []Value{NewInteger(85)})
	q := obj.DispatchPropertyGet("Quality")
	if q.Num != 85 {
		t.Errorf("expected Quality 85, got %d", q.Num)
	}

	// Test Interpolation get/set
	obj.DispatchPropertySet("Interpolation", []Value{NewInteger(3)})
	interp := obj.DispatchPropertyGet("Interpolation")
	if interp.Num != 3 {
		t.Errorf("expected Interpolation 3, got %d", interp.Num)
	}

	// Test resize via Width / Height setter
	obj.DispatchPropertySet("Width", []Value{NewInteger(100)})
	width := obj.DispatchPropertyGet("Width")
	if width.Num != 100 {
		t.Errorf("expected Width 100 after resize, got %d", width.Num)
	}
	height := obj.DispatchPropertyGet("Height")
	if height.Num != 150 {
		t.Errorf("expected Height to remain 150, got %d", height.Num)
	}

	obj.DispatchPropertySet("Height", []Value{NewInteger(75)})
	height = obj.DispatchPropertyGet("Height")
	if height.Num != 75 {
		t.Errorf("expected Height 75 after resize, got %d", height.Num)
	}

	// Test Canvas sub-object retrieval
	canvasVal := obj.DispatchPropertyGet("Canvas")
	if canvasVal.Type != VTNativeObject {
		t.Fatalf("expected Canvas to be VTNativeObject, got %v", canvasVal.Type)
	}

	canvasObj := vm.g3imageCanvasItems[canvasVal.Num]
	if canvasObj == nil {
		t.Fatal("expected canvasObj in VM map")
	}

	// Test Font retrieval from Canvas
	fontVal := vm.dispatchG3ImageCanvasPropertyGet(canvasObj, "Font")
	if fontVal.Type != VTNativeObject {
		t.Fatalf("expected Font to be VTNativeObject, got %v", fontVal.Type)
	}
	fontObj := vm.g3imageFontItems[fontVal.Num]
	if fontObj == nil {
		t.Fatal("expected fontObj in VM map")
	}

	// Test Font properties get/set
	vm.dispatchG3ImageFontPropertySet(fontObj, "Size", NewDouble(12.5))
	sizeVal := vm.dispatchG3ImageFontPropertyGet(fontObj, "Size")
	if sizeVal.Flt != 12.5 {
		t.Errorf("expected Font.Size 12.5, got %f", sizeVal.Flt)
	}

	vm.dispatchG3ImageFontPropertySet(fontObj, "Bold", NewBool(true))
	boldVal := vm.dispatchG3ImageFontPropertyGet(fontObj, "Bold")
	if boldVal.Type != VTBool || boldVal.Num == 0 {
		t.Error("expected Font.Bold true")
	}

	// Test Pen retrieval from Canvas
	penVal := vm.dispatchG3ImageCanvasPropertyGet(canvasObj, "Pen")
	if penVal.Type != VTNativeObject {
		t.Fatalf("expected Pen to be VTNativeObject, got %v", penVal.Type)
	}
	penObj := vm.g3imagePenItems[penVal.Num]
	if penObj == nil {
		t.Fatal("expected penObj in VM map")
	}

	// Test Pen properties get/set
	vm.dispatchG3ImagePenPropertySet(penObj, "Width", NewDouble(3.5))
	wVal := vm.dispatchG3ImagePenPropertyGet(penObj, "Width")
	if wVal.Flt != 3.5 {
		t.Errorf("expected Pen.Width 3.5, got %f", wVal.Flt)
	}

	// Test drawing operations
	vm.dispatchG3ImageCanvasMethod(canvasObj, "DrawLine", []Value{NewInteger(0), NewInteger(0), NewInteger(10), NewInteger(10)})
	vm.dispatchG3ImageCanvasMethod(canvasObj, "DrawBar", []Value{NewInteger(10), NewInteger(10), NewInteger(50), NewInteger(50)})

	// Test memory cleanup
	vm.cleanupG3ImageResources()
	if len(vm.g3imageItems) != 0 || len(vm.g3imageCanvasItems) != 0 || len(vm.g3imageFontItems) != 0 || len(vm.g3imagePenItems) != 0 {
		t.Error("expected all image sub-object items to be cleared by cleanupG3ImageResources")
	}
}
