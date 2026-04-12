/*
 * AxonASP Server
 * Copyright (C) 2026 G3pix Ltda. All rights reserved.
 *
 * Developed by Lucas GuimarÃ£es - G3pix Ltda
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

func TestG3PDF(t *testing.T) {
	vm := NewVM(nil, nil, 0)

	pdf := NewG3PDF(vm)
	if pdf == nil {
		t.Fatal("Failed to create G3PDF")
	}

	// basic interactions
	pdf.DispatchMethod("AddPage", nil)
	pdf.DispatchMethod("SetFont", []Value{NewString("Arial"), NewString("B"), NewInteger(16)})
	pdf.DispatchMethod("Cell", []Value{NewInteger(40), NewInteger(10), NewString("Hello World!")})

	page := pdf.DispatchPropertyGet("Page")
	if page.Num != 1 {
		t.Errorf("expected page 1, got %d", page.Num)
	}
}
