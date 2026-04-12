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

import "testing"

// TestVMRequestCollections verifies Request collection dispatch paths.
func TestVMRequestCollections(t *testing.T) {
	vm := NewVM(nil, nil, 5)
	host := NewMockHost()
	host.Request().QueryString.Add("q", "hello")
	host.Request().Form.Add("f", "world")
	host.Request().Cookies.AddCookie("profile", "name=lucas")
	vm.SetHost(host)

	query := vm.dispatchNativeCall(1, "QueryString", []Value{NewString("q")})
	if query.Type != VTString || query.Str != "hello" {
		t.Fatalf("unexpected QueryString result: %#v", query)
	}

	cookieSub := vm.dispatchNativeCall(1, "Cookies", []Value{NewString("profile"), NewString("name")})
	if cookieSub.Type != VTString || cookieSub.Str != "lucas" {
		t.Fatalf("unexpected cookie subkey result: %#v", cookieSub)
	}
}

// TestVMRequestDefaultCall verifies Request("key") style behavior through OpCall-native dispatch.
func TestVMRequestDefaultCall(t *testing.T) {
	vm := NewVM(nil, nil, 5)
	host := NewMockHost()
	host.Request().Form.Add("k", "value")
	vm.SetHost(host)

	result := vm.dispatchNativeCall(1, "", []Value{NewString("k")})
	if result.Type != VTString || result.Str != "value" {
		t.Fatalf("unexpected default Request lookup result: %#v", result)
	}
}

// TestVMRequestBinaryRead verifies Request.BinaryRead return behavior in VM dispatch.
func TestVMRequestBinaryRead(t *testing.T) {
	vm := NewVM(nil, nil, 5)
	host := NewMockHost()
	host.Request().SetBody([]byte("xyz"))
	vm.SetHost(host)

	value := vm.dispatchNativeCall(1, "BinaryRead", []Value{NewInteger(2)})
	if value.Type != VTString || value.Str != "xy" {
		t.Fatalf("unexpected binary read result: %#v", value)
	}

	left := vm.dispatchNativeCall(1, "BinaryRead", []Value{NewInteger(4)})
	if left.Type != VTString || left.Str != "z" {
		t.Fatalf("unexpected trailing binary read result: %#v", left)
	}
}
