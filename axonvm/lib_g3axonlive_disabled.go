//go:build lib_g3axonlive_disabled

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
 */
package axonvm

// G3AXONLIVE is the disabled stub for the G3AXONLIVE reactive component library.
type G3AXONLIVE struct{}

// newG3AxonLiveObject panics when the lib_g3axonlive_disabled build tag is set.
func (vm *VM) newG3AxonLiveObject() Value {
	panicLibraryDisabled("g3axonlive", "G3AXONLIVE library")
	return Value{Type: VTEmpty}
}

// DispatchPropertyGet is a no-op stub for the disabled library.
func (g *G3AXONLIVE) DispatchPropertyGet(_ string) Value {
	return Value{Type: VTEmpty}
}

// DispatchPropertySet is a no-op stub for the disabled library.
func (g *G3AXONLIVE) DispatchPropertySet(_ string, _ []Value) {}

// DispatchMethod is a no-op stub for the disabled library.
func (g *G3AXONLIVE) DispatchMethod(_ string, _ []Value) Value {
	return Value{Type: VTEmpty}
}

// ---------------------------------------------------------------------------
// Package-level stubs — required so server/fastcgi packages compile correctly
// when the lib_g3axonlive_disabled build tag is set.
// ---------------------------------------------------------------------------

// G3ALRegisterPage is a no-op stub for the disabled library.
func G3ALRegisterPage(_ string, _ string) {}

// G3ALGetPageForSession is a no-op stub; always returns empty string.
func G3ALGetPageForSession(_ string) string { return "" }

// G3ALStartCleanup is a no-op stub for the disabled library.
func G3ALStartCleanup(_ int) {}

// G3ALStopCleanup is a no-op stub for the disabled library.
func G3ALStopCleanup() {}
