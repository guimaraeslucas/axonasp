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
	"math"
	"testing"
)

// TestVMRequestCollections verifies Request collection dispatch paths.
func TestVMRequestCollections(t *testing.T) {
	vm := NewVM(nil, nil, 5)
	host := NewMockHost()
	host.Request().QueryString.Add("q", "hello")
	host.Request().Form.Add("f", "world")
	host.Request().Cookies.AddCookie("profile", "name=lucas")
	vm.SetHost(host)

	query := vm.dispatchNativeCall(1, "QueryString", []Value{NewString("q")})
	if query.Type != VTNativeObject {
		t.Fatalf("unexpected QueryString result: %#v", query)
	}
	queryValue := vm.dispatchNativeCall(query.Num, "", nil)
	if queryValue.Type != VTString || queryValue.Str != "hello" {
		t.Fatalf("unexpected QueryString resolved value: %#v", queryValue)
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
	if result.Type != VTNativeObject || vm.valueToString(result) != "value" {
		t.Fatalf("unexpected default Request lookup result: %#v", result)
	}
}

func TestVMRequestDefaultCallPreservesEmptyFormValue(t *testing.T) {
	vm := NewVM(nil, nil, 5)
	host := NewMockHost()
	host.Request().Form.Add("optional_field", "")
	host.Request().Cookies.AddCookie("optional_field", "cookie-fallback")
	vm.SetHost(host)

	result := vm.dispatchNativeCall(1, "", []Value{NewString("optional_field")})
	if result.Type != VTNativeObject {
		t.Fatalf("expected request collection value, got %#v", result)
	}
	if got := vm.valueToString(result); got != "" {
		t.Fatalf("expected present empty form value, got %q", got)
	}
	if count := vm.dispatchMemberGet(result, "Count"); count.Type != VTInteger || count.Num != 1 {
		t.Fatalf("unexpected request value Count: %#v", count)
	}
}

func TestVMRequestDefaultCallMissingValuePreservesCollectionSemantics(t *testing.T) {
	vm := NewVM(nil, nil, 5)
	vm.SetHost(NewMockHost())

	result := vm.dispatchNativeCall(nativeObjectRequest, "", []Value{NewString("missing")})
	if result.Type != VTNativeObject {
		t.Fatalf("expected missing Request value wrapper, got %#v", result)
	}
	if got := vm.jsToString(result); got != "undefined" {
		t.Fatalf("expected undefined string conversion, got %q", got)
	}
	if got := vm.jsToNumber(result); got.Type != VTDouble || !math.IsNaN(got.Flt) {
		t.Fatalf("expected missing Request value to convert to NaN, got %#v", got)
	}
}

func TestJScriptMissingRequestStringRemainsEmptyAcrossFunctionReturn(t *testing.T) {
	host := NewMockHost()
	source := `<%@ Language="JScript" %><%
function initialize() {
    var optionalValue = String(Request("optional_value"));
    Response.Write(typeof optionalValue + "|" + optionalValue + "|" + Server.URLEncode(optionalValue));
}
initialize();
%>`

	if got := runASPSourceForTestWithHost(t, source, host); got != "string|undefined|undefined" {
		t.Fatalf("unexpected missing Request value after function return: %q", got)
	}
}

func TestJScriptMissingRequestStringIsEmptyAtTopLevel(t *testing.T) {
	host := NewMockHost()
	source := `<%@ Language="JScript" %><% Response.Write("[" + String(Request("optional_value")) + "]"); %>`

	if got := runASPSourceForTestWithHost(t, source, host); got != "[undefined]" {
		t.Fatalf("unexpected top-level missing Request value: %q", got)
	}
}

func TestJScriptRequestDistinguishesMissingAndPresentEmptyValues(t *testing.T) {
	host := NewMockHost()
	host.Request().QueryString.SetLazyPayload([]byte("mode=new&optional_type="))
	source := `<%@ Language="JScript" %><%
var missing = String(Request("selection"));
var empty = String(Request("optional_type"));
var selected = missing == "undefined" ? "default" : "first";
Response.Write(missing + "|" + empty + "|" + selected);
%>`

	if got := runASPSourceForTestWithHost(t, source, host); got != "undefined||default" {
		t.Fatalf("unexpected missing/present-empty Request distinction: %q", got)
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
