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
	"os"
	"testing"
)

func TestG3Template(t *testing.T) {
	vm := NewVM(nil, nil, 0)
	vm.host = &MockHost{}

	tplLib := vm.newG3TemplateObject()
	if tplLib.Type != VTNativeObject {
		t.Fatalf("expected VTNativeObject, got %v", tplLib.Type)
	}

	obj := vm.g3templateItems[tplLib.Num]
	if obj == nil {
		t.Fatal("expected object in vm items")
	}

	// Create a dummy template file
	f, _ := os.Create("test.html")
	f.WriteString("Hello {{.name}}")
	f.Close()
	defer os.Remove("test.html")

	dictVal := vm.newDictionaryObject()
	vm.dispatchDictionaryMethod(dictVal.Num, "Add", []Value{NewString("name"), NewString("Axon")})

	res := obj.DispatchMethod("Render", []Value{NewString("test.html"), dictVal})
	if res.String() != "Hello Axon" {
		t.Errorf("expected Hello Axon, got %s", res.String())
	}
}
