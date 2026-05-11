//go:build lib_g3stringbuilder_disabled

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

// G3StringBuilder is the disabled stub for the G3STRINGBUILDER library.
type G3StringBuilder struct{}

// NewG3StringBuilder returns a runtime error because the library is disabled.
func NewG3StringBuilder() *G3StringBuilder {
	panicLibraryDisabled("g3stringbuilder", "G3STRINGBUILDER library")
	return nil
}

// DispatchMethod returns Empty when the library is disabled.
func (sb *G3StringBuilder) DispatchMethod(methodName string, args []Value) Value {
	return Value{Type: VTEmpty}
}

// DispatchPropertyGet returns Empty when the library is disabled.
func (sb *G3StringBuilder) DispatchPropertyGet(propertyName string) Value {
	return Value{Type: VTEmpty}
}

// DispatchPropertySet does nothing when the library is disabled.
func (sb *G3StringBuilder) DispatchPropertySet(propertyName string, val Value) {}
