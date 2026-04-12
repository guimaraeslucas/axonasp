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

func TestG3JSON(t *testing.T) {
	vm := NewVM(nil, nil, 0)
	vm.host = &MockHost{}

	jsonLib := vm.newG3JSONObject()
	if jsonLib.Type != VTNativeObject {
		t.Fatalf("expected VTNativeObject, got %v", jsonLib.Type)
	}

	obj := vm.g3jsonItems[jsonLib.Num]
	if obj == nil {
		t.Fatal("expected object in vm items")
	}

	// Test Parse String
	parsed := obj.DispatchMethod("Parse", []Value{NewString(`{"name":"test", "age":30}`)})
	if parsed.Type != VTNativeObject {
		t.Fatalf("expected parsed to be dictionary, got %v", parsed.Type)
	}

	dict := vm.dictionaryItems[parsed.Num]
	if dict == nil {
		t.Fatal("dictionary missing")
	}

	nameVal, _ := vm.dispatchDictionaryPropertyGet(parsed.Num, "Item") // Wait, Item takes args. It's actually a method call in VBScript if parameterized. Let's just use method dispatch.
	nameVal, _ = vm.dispatchDictionaryMethod(parsed.Num, "Item", []Value{NewString("name")})
	if nameVal.String() != "test" {
		t.Errorf("expected test, got %s", nameVal.String())
	}

	// Test Stringify
	str := obj.DispatchMethod("Stringify", []Value{parsed})
	if str.Type != VTString {
		t.Fatalf("expected VTString, got %v", str.Type)
	}
}
