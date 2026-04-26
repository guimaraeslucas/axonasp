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
	"strings"
	"testing"
)

// TestVMServerCreateObjectG3Test verifies native G3TestSuite activation through Server.CreateObject.
func TestVMServerCreateObjectG3Test(t *testing.T) {
	vm := NewVM(nil, nil, 16)
	vm.host = NewMockHost()

	obj := vm.dispatchNativeCall(nativeObjectServer, "CreateObject", []Value{NewString("G3TestSuite")})
	if obj.Type != VTNativeObject {
		t.Fatalf("expected VTNativeObject, got %#v", obj)
	}
	if _, ok := vm.g3testItems[obj.Num]; !ok {
		t.Fatalf("expected G3TestSuite object stored in VM map")
	}

	alias := vm.dispatchNativeCall(nativeObjectServer, "CreateObject", []Value{NewString("G3Test")})
	if alias.Type != VTNativeObject {
		t.Fatalf("expected VTNativeObject alias, got %#v", alias)
	}
}

// TestG3TestDispatchAndSummary verifies assertion accounting and summary extraction.
func TestG3TestDispatchAndSummary(t *testing.T) {
	vm := NewVM(nil, nil, 16)
	vm.host = NewMockHost()
	obj := vm.dispatchNativeCall(nativeObjectServer, "CreateObject", []Value{NewString("G3TestSuite")})
	if obj.Type != VTNativeObject {
		t.Fatalf("expected VTNativeObject, got %#v", obj)
	}

	vm.dispatchNativeCall(obj.Num, "Describe", []Value{NewString("math block")})
	if res := vm.dispatchNativeCall(obj.Num, "AssertEqual", []Value{NewInteger(10), NewDouble(10), NewString("numeric coercion")}); !vm.asBool(res) {
		t.Fatalf("expected numeric equality to pass, got %#v", res)
	}
	if res := vm.dispatchNativeCall(obj.Num, "AssertTrue", []Value{NewInteger(1), NewString("truthy integer")}); !vm.asBool(res) {
		t.Fatalf("expected truthy assertion to pass, got %#v", res)
	}
	if res := vm.dispatchNativeCall(obj.Num, "AssertFalse", []Value{NewBool(false), NewString("false branch")}); !vm.asBool(res) {
		t.Fatalf("expected false assertion to pass, got %#v", res)
	}
	if res := vm.dispatchNativeCall(obj.Num, "AssertNotEqual", []Value{NewInteger(10), NewInteger(11), NewString("values should differ")}); !vm.asBool(res) {
		t.Fatalf("expected inequality assertion to pass, got %#v", res)
	}
	if res := vm.dispatchNativeCall(obj.Num, "AssertEmpty", []Value{NewEmpty(), NewString("uninitialized values should be Empty")}); !vm.asBool(res) {
		t.Fatalf("expected empty assertion to pass, got %#v", res)
	}
	if res := vm.dispatchNativeCall(obj.Num, "AssertNull", []Value{{Type: VTNull}, NewString("explicit Null should pass")}); !vm.asBool(res) {
		t.Fatalf("expected null assertion to pass, got %#v", res)
	}
	if res := vm.dispatchNativeCall(obj.Num, "AssertNothing", []Value{Value{Type: VTObject, Num: 0}, NewString("Nothing object references should pass")}); !vm.asBool(res) {
		t.Fatalf("expected nothing assertion to pass, got %#v", res)
	}
	if res := vm.dispatchNativeCall(obj.Num, "AssertTypeName", []Value{NewString("String"), NewString("abc"), NewString("TypeName(String) should match")}); !vm.asBool(res) {
		t.Fatalf("expected typename assertion to pass, got %#v", res)
	}
	if res := vm.dispatchNativeCall(obj.Num, "AssertLength", []Value{NewInteger(3), NewString("abc"), NewString("string length should be 3")}); !vm.asBool(res) {
		t.Fatalf("expected length assertion to pass, got %#v", res)
	}
	if res := vm.dispatchNativeCall(obj.Num, "AssertCount", []Value{NewInteger(2), Value{Type: VTArray, Arr: NewVBArrayFromValues(0, []Value{NewInteger(1), NewInteger(2)})}, NewString("array count should be 2")}); !vm.asBool(res) {
		t.Fatalf("expected count alias assertion to pass, got %#v", res)
	}
	if res := vm.dispatchNativeCall(obj.Num, "AssertRaises", []Value{NewString("Err.Raise 13, \"suite\", \"type mismatch\""), NewInteger(13), NewString("Err.Raise should surface explicit numbers")}); !vm.asBool(res) {
		t.Fatalf("expected AssertRaises numeric check to pass, got %#v", res)
	}
	if res := vm.dispatchNativeCall(obj.Num, "AssertRaises", []Value{NewString("Err.Raise 13, \"suite\", \"type mismatch\""), NewString("type mismatch"), NewString("AssertRaises should match description text")}); !vm.asBool(res) {
		t.Fatalf("expected AssertRaises text check to pass, got %#v", res)
	}
	if res := vm.dispatchNativeCall(obj.Num, "AssertRaises", []Value{NewString("Function Broken("), NewString("syntax error should be trapped")}); !vm.asBool(res) {
		t.Fatalf("expected AssertRaises syntax trap to pass, got %#v", res)
	}
	if res := vm.dispatchNativeCall(obj.Num, "AssertEqual", []Value{NewString("abc"), NewString("xyz"), NewString("string mismatch")}); vm.asBool(res) {
		t.Fatalf("expected mismatch assertion to fail")
	}
	vm.dispatchNativeCall(obj.Num, "Fail", []Value{NewString("forced failure")})

	reports := vm.GetG3TestReports()
	if len(reports) != 1 {
		t.Fatalf("expected 1 report, got %d", len(reports))
	}
	if reports[0].SuiteName != "math block" {
		t.Fatalf("expected suite name to be preserved, got %q", reports[0].SuiteName)
	}

	summary := vm.GetG3TestSuiteSummary()
	if summary.SuiteCount != 1 {
		t.Fatalf("expected suite count 1, got %d", summary.SuiteCount)
	}
	if summary.Total != 15 || summary.Passed != 13 || summary.Failed != 2 {
		t.Fatalf("unexpected summary counters: %+v", summary)
	}
	if len(summary.Failures) != 2 {
		t.Fatalf("expected 2 failures, got %d", len(summary.Failures))
	}
	if !strings.Contains(summary.Failures[0].Message, "math block") {
		t.Fatalf("expected suite prefix in failure message, got %q", summary.Failures[0].Message)
	}
	if !strings.Contains(summary.Failures[0].Message, "expected=abc actual=xyz") {
		t.Fatalf("expected diff content in failure message, got %q", summary.Failures[0].Message)
	}
	if !strings.Contains(summary.Failures[1].Message, "forced failure") {
		t.Fatalf("expected explicit failure message, got %q", summary.Failures[1].Message)
	}
}

// TestG3TestLegacySuiteMethods verifies G3TestSuite compatibility methods used by legacy ASP fixtures.
func TestG3TestLegacySuiteMethods(t *testing.T) {
	vm := NewVM(nil, nil, 16)
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	host.Response().SetBuffer(false)
	vm.host = host

	obj := vm.dispatchNativeCall(nativeObjectServer, "CreateObject", []Value{NewString("G3TestSuite")})
	if obj.Type != VTNativeObject {
		t.Fatalf("expected VTNativeObject, got %#v", obj)
	}

	vm.dispatchNativeCall(obj.Num, "BeginTest", []Value{NewString("legacy loop suite")})
	vm.dispatchNativeCall(obj.Num, "SetVar", []Value{NewString("whileResult"), NewString("0,1,2,")})
	got := vm.dispatchNativeCall(obj.Num, "GetVar", []Value{NewString("whileResult")})
	if got.String() != "0,1,2," {
		t.Fatalf("unexpected SetVar/GetVar roundtrip: %q", got.String())
	}

	if res := vm.dispatchNativeCall(obj.Num, "AssertEquals", []Value{NewString("0,1,2,"), got, NewString("legacy while output")}); !vm.asBool(res) {
		t.Fatalf("expected assertion pass, got %#v", res)
	}
	if res := vm.dispatchNativeCall(obj.Num, "AssertNotEquals", []Value{NewString("0,1,2,3,"), got, NewString("legacy mismatch alias")}); !vm.asBool(res) {
		t.Fatalf("expected inequality alias to pass, got %#v", res)
	}
	vm.dispatchNativeCall(obj.Num, "EndTest", nil)
	vm.dispatchNativeCall(obj.Num, "Summary", nil)

	text := output.String()
	if !strings.Contains(text, "G3Test Summary: total=2, passed=2, failed=0") {
		t.Fatalf("expected summary output, got %q", text)
	}
}
