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
