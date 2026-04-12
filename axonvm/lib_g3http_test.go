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

func TestG3HTTP(t *testing.T) {
	vm := NewVM(nil, nil, 0)

	httpLib := vm.newG3HTTPObject()
	if httpLib.Type != VTNativeObject {
		t.Fatalf("expected VTNativeObject, got %v", httpLib.Type)
	}

	// Normally we would mock HTTP, but just verifying the dispatch method exists
	obj := vm.g3httpItems[httpLib.Num]
	if obj == nil {
		t.Fatal("expected object in vm items")
	}

	// Invalid fetch should return empty or error string rather than crash
	res := obj.DispatchMethod("Fetch", []Value{NewString("invalid-url")})
	if res.Type != VTString && res.Type != VTEmpty {
		t.Fatalf("unexpected type %v", res.Type)
	}
}
