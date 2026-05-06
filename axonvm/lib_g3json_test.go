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
	"bytes"
	"encoding/json"
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

// TestG3JSONStringifyVBScriptPatchShape verifies the exact VBScript pattern used
// by AxonLive pages: NewObject + NewArray + ReDim Preserve + Set array(index)=object.
func TestG3JSONStringifyVBScriptPatchShape(t *testing.T) {
	source := `<%
Dim j, responseObj, components, patch
Set j = Server.CreateObject("G3JSON")

Set responseObj = j.NewObject()
responseObj("success") = True

components = j.NewArray()
ReDim Preserve components(0)
Set patch = j.NewObject()
patch("componentId") = "lblCounter"
patch("HTML") = "<div>Counter: 1</div>"
Set components(0) = patch

responseObj("components") = components
Response.Write j.Stringify(responseObj)
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	if output.Len() == 0 {
		t.Fatal("expected non-empty JSON output")
	}

	var parsed struct {
		Success    bool `json:"success"`
		Components []struct {
			ComponentID string `json:"componentId"`
			HTML        string `json:"HTML"`
		} `json:"components"`
	}
	if err := json.Unmarshal(output.Bytes(), &parsed); err != nil {
		t.Fatalf("invalid JSON output %q: %v", output.String(), err)
	}

	if !parsed.Success {
		t.Fatalf("expected success=true, got false in %q", output.String())
	}
	if len(parsed.Components) != 1 {
		t.Fatalf("expected 1 component patch, got %d in %q", len(parsed.Components), output.String())
	}
	if parsed.Components[0].ComponentID != "lblCounter" {
		t.Fatalf("unexpected componentId: %q", parsed.Components[0].ComponentID)
	}
	if parsed.Components[0].HTML != "<div>Counter: 1</div>" {
		t.Fatalf("unexpected HTML payload: %q", parsed.Components[0].HTML)
	}
}

// TestG3JSONStringifyNestedArraysOfObjects verifies nested array/object payloads
// serialize correctly without requiring manual JSON writing in ASP.
func TestG3JSONStringifyNestedArraysOfObjects(t *testing.T) {
	source := `<%
Dim j, responseObj, components, patch, childRows, child
Set j = Server.CreateObject("G3JSON")

Set responseObj = j.NewObject()
responseObj("success") = True

components = j.NewArray()
ReDim Preserve components(0)

Set patch = j.NewObject()
patch("componentId") = "panelRoot"

childRows = j.NewArray()
ReDim Preserve childRows(1)
Set child = j.NewObject()
child("name") = "first"
child("value") = 1
Set childRows(0) = child

Set child = j.NewObject()
child("name") = "second"
child("value") = 2
Set childRows(1) = child

patch("rows") = childRows
Set components(0) = patch

responseObj("components") = components
Response.Write j.Stringify(responseObj)
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	if output.Len() == 0 {
		t.Fatal("expected non-empty JSON output")
	}

	var parsed map[string]interface{}
	if err := json.Unmarshal(output.Bytes(), &parsed); err != nil {
		t.Fatalf("invalid JSON output %q: %v", output.String(), err)
	}

	componentsVal, ok := parsed["components"].([]interface{})
	if !ok || len(componentsVal) != 1 {
		t.Fatalf("expected components array with one element, got %#v", parsed["components"])
	}

	patchObj, ok := componentsVal[0].(map[string]interface{})
	if !ok {
		t.Fatalf("expected patch object, got %#v", componentsVal[0])
	}

	rows, ok := patchObj["rows"].([]interface{})
	if !ok || len(rows) != 2 {
		t.Fatalf("expected rows array with two objects, got %#v", patchObj["rows"])
	}
}

// TestG3JSONStringifyTwoPatchComponents mirrors AxonLive async responses where
// two components are patched in one cycle (label + button with quoted attributes).
func TestG3JSONStringifyTwoPatchComponents(t *testing.T) {
	source := `<%
Dim j, responseObj, components, patch
Set j = Server.CreateObject("G3JSON")

Set responseObj = j.NewObject()
responseObj("success") = True

components = j.NewArray()

ReDim Preserve components(0)
Set patch = j.NewObject()
patch("componentId") = "lblCounter"
patch("HTML") = "<div id=""lblCounter"">Counter: 1</div>"
Set components(0) = patch

ReDim Preserve components(1)
Set patch = j.NewObject()
patch("componentId") = "btnIncrement"
patch("HTML") = "<button id=""btnIncrement"" data-g3al-id=""btnIncrement"">Increment Counter</button>"
Set components(1) = patch

responseObj("components") = components
Response.Write j.Stringify(responseObj)
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	if output.Len() == 0 {
		t.Fatal("expected non-empty JSON output")
	}

	var parsed struct {
		Success    bool `json:"success"`
		Components []struct {
			ComponentID string `json:"componentId"`
			HTML        string `json:"HTML"`
		} `json:"components"`
	}
	if err := json.Unmarshal(output.Bytes(), &parsed); err != nil {
		t.Fatalf("invalid JSON output %q: %v", output.String(), err)
	}

	if !parsed.Success {
		t.Fatalf("expected success=true, got false in %q", output.String())
	}
	if len(parsed.Components) != 2 {
		t.Fatalf("expected 2 component patches, got %d in %q", len(parsed.Components), output.String())
	}
}
