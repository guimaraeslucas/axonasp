//go:build wasm || lib_g3image_disabled

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

// G3Image is the disabled stub for the G3Image library.
type G3Image struct{}

type G3ImageCanvas struct{}
type G3ImageFont struct{}
type G3ImagePen struct{}

func (vm *VM) newG3ImageObject() Value {
	panicLibraryDisabled("g3image", "G3Image library")
	return Value{Type: VTEmpty}
}

func (vm *VM) newPersitsJpegObject() Value {
	panicLibraryDisabled("g3image", "G3Image library")
	return Value{Type: VTEmpty}
}

func (g *G3Image) DispatchPropertyGet(propertyName string) Value {
	return Value{Type: VTEmpty}
}

func (g *G3Image) DispatchPropertySet(propertyName string, args []Value) bool {
	return false
}

func (g *G3Image) DispatchMethod(methodName string, args []Value) Value {
	return Value{Type: VTEmpty}
}

func (vm *VM) cleanupG3ImageResources() {}

func (vm *VM) dispatchG3ImageCanvasMethod(c *G3ImageCanvas, member string, args []Value) Value {
	return Value{Type: VTEmpty}
}

func (vm *VM) dispatchG3ImageCanvasPropertyGet(c *G3ImageCanvas, member string) Value {
	return Value{Type: VTEmpty}
}

func (vm *VM) dispatchG3ImageFontPropertyGet(f *G3ImageFont, member string) Value {
	return Value{Type: VTEmpty}
}

func (vm *VM) dispatchG3ImageFontPropertySet(f *G3ImageFont, member string, val Value) bool {
	return false
}

func (vm *VM) dispatchG3ImagePenPropertyGet(p *G3ImagePen, member string) Value {
	return Value{Type: VTEmpty}
}

func (vm *VM) dispatchG3ImagePenPropertySet(p *G3ImagePen, member string, val Value) bool {
	return false
}
