//go:build lib_g3md_disabled

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
	"fmt"
	"g3pix.com.br/axonasp/v2/vbscript"
)

// G3MD é a estrutura para a versão desativada da library.
type G3MD struct{}

func NewG3MD() *G3MD {
	panic(&VMError{
		Code:        vbscript.ActiveXCannotCreateObject,
		Number:      int(ErrLibraryDisabled),
		Description: fmt.Sprintf(ErrLibraryDisabled.String(), "g3md"),
		Source:      "G3MD library",
	})
}

// DispatchMethod implements the method dispatching for G3MD. Since the library is disabled, it will return empty values or do nothing.
func (md *G3MD) DispatchMethod(methodName string, args []Value) Value {
	return Value{Type: VTEmpty}
}

// DispatchPropertyGet implements the property get dispatching for G3MD. Since the library is disabled, it will return empty values or do nothing.
func (md *G3MD) DispatchPropertyGet(propertyName string) Value {
	return Value{Type: VTEmpty}
}

// DispatchPropertySet implements the property set dispatching for G3MD. Since the library is disabled, it will return empty values or do nothing.
func (md *G3MD) DispatchPropertySet(propertyName string, val Value) {
}
