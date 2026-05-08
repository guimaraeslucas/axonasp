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

// TestG3AxonLiveGranularMutations verifies the new object-oriented API for component manipulation.
func TestG3AxonLiveGranularMutations(t *testing.T) {
	vm := NewVM(nil, nil, 10)
	host := NewMockHost()
	vm.SetHost(host)

	// 1. Create G3AXONLIVE object.
	axonLiveVal := vm.dispatchNativeCall(nativeObjectServer, "CreateObject", []Value{NewString("G3AXONLIVE")})
	if axonLiveVal.Type != VTNativeObject {
		t.Fatalf("expected G3AXONLIVE object, got %#v", axonLiveVal)
	}
	axonLiveID := axonLiveVal.Num
	axonLive := vm.g3axonliveItems[axonLiveID]

	// 2. Initialize page (normal load).
	vm.dispatchNativeCall(axonLiveID, "InitPage", nil)
	sessionID := "test-session"
	host.Session().ID = sessionID

	// 3. Get a component proxy.
	btnProxyVal := vm.dispatchNativeCall(axonLiveID, "GetComponent", []Value{NewString("btnSubmit")})
	if btnProxyVal.Type != VTNativeObject {
		t.Fatalf("expected component proxy object, got %#v", btnProxyVal)
	}
	btnID := btnProxyVal.Num

	// 4. Test Property Set (set_property action).
	vm.dispatchMemberSet(btnID, "value", NewString("Lucas"))

	foundSetProperty := false
	for _, action := range axonLive.pendingActions {
		if action.Type == "set_property" && action.ComponentID == "btnSubmit" && action.AttrName == "value" && action.AttrValue == "Lucas" {
			foundSetProperty = true
			break
		}
	}
	if !foundSetProperty {
		t.Error("set_property action not found in pendingActions")
	}

	// 5. Test Persistence (Property Get).
	val := vm.dispatchMemberGet(btnProxyVal, "value")
	if val.Type != VTString || val.Str != "Lucas" {
		t.Errorf("expected persistent value 'Lucas', got %#v", val)
	}

	// 6. Test Methods (SetStyle, AddClass, etc.)
	vm.dispatchNativeCall(btnID, "SetStyle", []Value{NewString("background-color"), NewString("red")})
	vm.dispatchNativeCall(btnID, "AddClass", []Value{NewString("highlight")})
	vm.dispatchNativeCall(btnID, "RemoveClass", []Value{NewString("old-class")})
	vm.dispatchNativeCall(btnID, "SetAttribute", []Value{NewString("data-test"), NewString("123")})
	vm.dispatchNativeCall(btnID, "RemoveAttribute", []Value{NewString("disabled")})
	vm.dispatchNativeCall(btnID, "AddTitle", []Value{NewString("Click me")})
	vm.dispatchNativeCall(btnID, "RemoveTitle", nil)
	vm.dispatchNativeCall(btnID, "SetValue", []Value{NewString("New Value")})

	expectedActions := []string{
		"set_style", "add_class", "remove_class", "add_attribute",
		"remove_attribute", "add_title", "remove_title", "set_value",
	}

	for _, expType := range expectedActions {
		found := false
		for _, action := range axonLive.pendingActions {
			if action.Type == expType && action.ComponentID == "btnSubmit" {
				found = true
				break
			}
		}
		if !found {
			t.Errorf("action type '%s' not found in pendingActions", expType)
		}
	}
}
