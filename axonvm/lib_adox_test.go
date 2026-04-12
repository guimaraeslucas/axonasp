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

func TestADOX(t *testing.T) {
	vm := NewVM(nil, nil, 0)

	adoxLib := vm.newADOXCatalogObject()
	if adoxLib.Type != VTNativeObject {
		t.Fatalf("expected VTNativeObject, got %v", adoxLib.Type)
	}

	obj := vm.adoxCatalogItems[adoxLib.Num]
	if obj == nil {
		t.Fatal("expected ADOXCatalog object")
	}

	// Just check if setting ActiveConnection works without crashing
	obj.DispatchPropertySet("ActiveConnection", []Value{NewString("Provider=SQLOLEDB;...")})
	conn := obj.DispatchPropertyGet("ActiveConnection")
	if conn.Type != VTString {
		t.Errorf("expected VTString connection, got %v", conn.Type)
	}
}

func TestADOXTablesEnumerationForEachBridge(t *testing.T) {
	vm := NewVM(nil, nil, 0)

	tablesObj := vm.newADOXTablesObject([]*ADOXTable{
		{Name: "Customers", Type: "TABLE"},
		{Name: "Orders", Type: "TABLE"},
	})

	enumResult, err := vbsAxonEnumValues(vm, []Value{tablesObj})
	if err != nil {
		t.Fatalf("vbsAxonEnumValues returned error: %v", err)
	}
	if enumResult.Type != VTArray || enumResult.Arr == nil {
		t.Fatalf("expected VTArray result, got %#v", enumResult)
	}
	if len(enumResult.Arr.Values) != 2 {
		t.Fatalf("expected 2 enumerated table objects, got %d", len(enumResult.Arr.Values))
	}

	first := enumResult.Arr.Values[0]
	if first.Type != VTNativeObject {
		t.Fatalf("expected first enumerated item to be VTNativeObject, got %#v", first)
	}
	firstName := vm.dispatchMemberGet(first, "Name")
	if firstName.Type != VTString || firstName.String() != "Customers" {
		t.Fatalf("expected first table name Customers, got %#v", firstName)
	}

	second := enumResult.Arr.Values[1]
	if second.Type != VTNativeObject {
		t.Fatalf("expected second enumerated item to be VTNativeObject, got %#v", second)
	}
	secondName := vm.dispatchMemberGet(second, "Name")
	if secondName.Type != VTString || secondName.String() != "Orders" {
		t.Fatalf("expected second table name Orders, got %#v", secondName)
	}
}
