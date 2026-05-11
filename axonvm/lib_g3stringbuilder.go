//go:build !lib_g3stringbuilder_disabled

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

import (
	"strings"
)

// G3StringBuilder wraps strings.Builder for low-allocation string accumulation.
type G3StringBuilder struct {
	builder strings.Builder
}

// NewG3StringBuilder creates a new G3STRINGBUILDER object.
func NewG3StringBuilder() *G3StringBuilder {
	return &G3StringBuilder{}
}

// DispatchMethod routes method calls for the G3STRINGBUILDER object.
func (sb *G3StringBuilder) DispatchMethod(methodName string, args []Value) Value {
	switch {
	case strings.EqualFold(methodName, "Append"):
		if len(args) == 0 {
			return Value{Type: VTEmpty}
		}
		sb.builder.WriteString(args[0].String())
		return Value{Type: VTEmpty}
	case strings.EqualFold(methodName, "ToString"):
		return NewString(sb.builder.String())
	default:
		return Value{Type: VTEmpty}
	}
}

// DispatchPropertyGet resolves read access for pseudo-property style usage.
func (sb *G3StringBuilder) DispatchPropertyGet(propertyName string) Value {
	switch {
	case strings.EqualFold(propertyName, "ToString"):
		return NewString(sb.builder.String())
	default:
		return Value{Type: VTEmpty}
	}
}

// DispatchPropertySet ignores property writes because this object is method-driven.
func (sb *G3StringBuilder) DispatchPropertySet(propertyName string, val Value) {
}
