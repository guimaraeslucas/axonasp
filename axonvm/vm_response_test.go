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
	"testing"
)

// TestVMResponseWriteAndProperties verifies Response dispatch for write and properties.
func TestVMResponseWriteAndProperties(t *testing.T) {
	vm := NewVM(nil, nil, 5)
	host := NewMockHost()
	vm.SetHost(host)

	vm.dispatchNativeCall(nativeObjectResponse, "ContentType", []Value{NewString("text/plain")})
	contentType := vm.dispatchNativeCall(nativeObjectResponse, "ContentType", nil)
	if contentType.Type != VTString || contentType.Str != "text/plain" {
		t.Fatalf("unexpected content type: %#v", contentType)
	}

	vm.dispatchNativeCall(nativeObjectResponse, "Write", []Value{NewString("ok")})
	if host.Response().GetContentType() != "text/plain" {
		t.Fatalf("unexpected response content type state")
	}
}

// TestVMResponseCookies verifies cookie dispatch methods.
func TestVMResponseCookies(t *testing.T) {
	vm := NewVM(nil, nil, 5)
	host := NewMockHost()
	vm.SetHost(host)

	vm.dispatchNativeCall(nativeObjectResponse, "Cookies", []Value{NewString("sid"), NewString("123")})
	cookie := vm.dispatchNativeCall(nativeObjectResponse, "Cookies", []Value{NewString("sid")})
	if cookie.Type != VTNativeObject {
		t.Fatalf("unexpected cookie value: %#v", cookie)
	}

	vm.dispatchNativeCall(nativeObjectResponse, "Cookies", []Value{NewString("sid"), NewString("Domain"), NewString("example.com")})
	domain := vm.dispatchNativeCall(nativeObjectResponse, "Cookies", []Value{NewString("sid"), NewString("Domain")})
	if domain.Type != VTString || domain.Str != "example.com" {
		t.Fatalf("unexpected cookie domain: %#v", domain)
	}
}

// TestVMResponseCookiesExpression verifies direct and chained cookie access through compiled expressions.
func TestVMResponseCookiesExpression(t *testing.T) {
	source := `<% Response.Cookies "sid", "abc" : Response.Cookies "sid", "Domain", "example.com" %><%= Response.Cookies("sid") %>|<%= Response.Cookies("sid").Domain %>`
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

	if output.String() != "abc|example.com" {
		t.Fatalf("unexpected output: %q", output.String())
	}
}

// TestVMResponseEndSignal verifies Response.End is treated as normal script termination.
func TestVMResponseEndSignal(t *testing.T) {
	source := `<% Response.Write "a" : Response.End : Response.Write "b" %>`
	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("expected no error for Response.End signal, got: %v", err)
	}
}

// TestVMResponseEndSignalPropagatesAcrossExecuteGlobal verifies Response.End in
// dynamic ExecuteGlobal code still terminates the outer page execution.
func TestVMResponseEndSignalPropagatesAcrossExecuteGlobal(t *testing.T) {
	source := `<%
ExecuteGlobal "Response.Write ""x"" : Response.End"
Response.Write "y"
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
		t.Fatalf("expected no error for ExecuteGlobal Response.End signal, got: %v", err)
	}
	if output.String() != "x" {
		t.Fatalf("expected only dynamic output before Response.End, got %q", output.String())
	}
}

// TestVMResponsePropertyGettersInExpression verifies that Response.Prop() in expression context
// returns the property value instead of Empty. Previously, OpMemberGet+OpCall(0) on a string/bool
// value would return Empty because OpCall only handles NativeObject/Builtin targets.
func TestVMResponsePropertyGettersInExpression(t *testing.T) {
	source := `<%
Response.ContentType "text/plain"
Response.CacheControl "No-Cache"
Response.Status "200 OK"
%><%= Response.ContentType() %>|<%= Response.CacheControl() %>|<%= Response.Status() %>|<%= Response.IsClientConnected() %>`
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

	// IsClientConnected returns True when no real HTTP request is set (mock host)
	expected := "text/plain|No-Cache|200 OK|True"
	if output.String() != expected {
		t.Fatalf("unexpected getter output:\nexpected: %q\nactual:   %q", expected, output.String())
	}
}

// TestVMResponseCookiesMultiArgExpression verifies Response.Cookies("name","Prop") in expression context
// returns the cookie property rather than the cookie value.
func TestVMResponseCookiesMultiArgExpression(t *testing.T) {
	source := `<% Response.Cookies "sid", "abc123" : Response.Cookies "sid", "Domain", "localhost" %><%= Response.Cookies("sid") %>|<%= Response.Cookies("sid", "Domain") %>|<%= Response.Cookies("sid").Domain %>`
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

	expected := "abc123|localhost|localhost"
	if output.String() != expected {
		t.Fatalf("unexpected cookie expression output:\nexpected: %q\nactual:   %q", expected, output.String())
	}
}

// TestVMResponseCookiesCountAndKey verifies Response.Cookies exposes Count and Key(index).
func TestVMResponseCookiesCountAndKey(t *testing.T) {
	source := `<%
Response.Cookies("Sid") = "1"
Response.Cookies("Token") = "2"
%><%= Response.Cookies.Count %>|<%= Response.Cookies.Key(1) %>|<%= Response.Cookies.Key(2) %>`

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

	if output.String() != "2|Sid|Token" {
		t.Fatalf("unexpected response cookies metadata output: %q", output.String())
	}
}

// TestVMResponseCookiesChainedPropertyAssignment verifies that
// Response.Cookies("name").Property = value compiles and executes correctly.
// This regression test covers the chained member assignment pattern in statement
// context (obj.Method(args).Prop = value) which previously failed with
// "Variable not defined: 'Expires'" due to the compiler discarding the call
// result before parsing the chained property set.
func TestVMResponseCookiesChainedPropertyAssignment(t *testing.T) {
	tests := []struct {
		name     string
		source   string
		expected string
	}{
		{
			name:     "expires property set",
			source:   `<% Response.Cookies("ck") = "val" : Response.Cookies("ck").Expires = "Wed, 21 Oct 2015 07:28:00 GMT" %><%= Response.Cookies("ck", "Expires") %>`,
			expected: "Wed, 21 Oct 2015 07:28:00 GMT",
		},
		{
			name:     "domain property set",
			source:   `<% Response.Cookies("ck") = "val" : Response.Cookies("ck").Domain = "example.com" %><%= Response.Cookies("ck", "Domain") %>`,
			expected: "example.com",
		},
		{
			name:     "path property set",
			source:   `<% Response.Cookies("ck") = "val" : Response.Cookies("ck").Path = "/app" %><%= Response.Cookies("ck", "Path") %>`,
			expected: "/app",
		},
		{
			name:     "secure property set",
			source:   `<% Response.Cookies("ck") = "val" : Response.Cookies("ck").Secure = True %><%= Response.Cookies("ck", "Secure") %>`,
			expected: "True",
		},
		{
			name:     "httponly property set",
			source:   `<% Response.Cookies("ck") = "val" : Response.Cookies("ck").HttpOnly = True %><%= Response.Cookies("ck", "HttpOnly") %>`,
			expected: "True",
		},
		{
			name:     "chained after two-arg call",
			source:   `<% Response.Cookies("ck", "val1") : Response.Cookies("ck").Path = "/test" %><%= Response.Cookies("ck", "Path") %>`,
			expected: "/test",
		},
		{
			name:     "multiple chained properties",
			source:   `<% Response.Cookies("ck") = "val" : Response.Cookies("ck").Domain = "test.com" : Response.Cookies("ck").Path = "/x" : Response.Cookies("ck").Secure = True %><%= Response.Cookies("ck", "Domain") %>|<%= Response.Cookies("ck", "Path") %>|<%= Response.Cookies("ck", "Secure") %>`,
			expected: "test.com|/x|True",
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			compiler := NewASPCompiler(tt.source)
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

			if output.String() != tt.expected {
				t.Fatalf("unexpected output:\nexpected: %q\nactual:   %q", tt.expected, output.String())
			}
		})
	}
}

// TestVMResponseBinaryWritePreservesVBByteString verifies BinaryWrite preserves raw VB byte-string payloads.
func TestVMResponseBinaryWritePreservesVBByteString(t *testing.T) {
	vm := NewVM(nil, nil, 5)
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	payload := bytesToVBByteString([]byte{'B', 'M', 0x80, 0xFF})
	vm.dispatchNativeCall(nativeObjectResponse, "BinaryWrite", []Value{NewString(payload)})
	host.Response().Flush()

	expected := []byte{'B', 'M', 0x80, 0xFF}
	if !bytes.Equal(output.Bytes(), expected) {
		t.Fatalf("unexpected BinaryWrite bytes: %v", output.Bytes())
	}
}

// TestVMResponseBinaryWriteSupportsVTArray verifies BinaryWrite accepts VTArray byte payloads.
func TestVMResponseBinaryWriteSupportsVTArray(t *testing.T) {
	vm := NewVM(nil, nil, 5)
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	payload := Value{Type: VTArray, Arr: NewVBArrayFromValues(0, []Value{NewInteger(0x89), NewInteger(0x50), NewInteger(0x4E), NewInteger(0x47)})}
	vm.dispatchNativeCall(nativeObjectResponse, "BinaryWrite", []Value{payload})
	host.Response().Flush()

	expected := []byte{0x89, 0x50, 0x4E, 0x47}
	if !bytes.Equal(output.Bytes(), expected) {
		t.Fatalf("unexpected BinaryWrite VTArray bytes: %v", output.Bytes())
	}
}

// TestVMResponseBinaryWriteSuppressesFormattingWhitespace verifies formatting-only gaps between ASP blocks do not leak into binary responses.
func TestVMResponseBinaryWriteSuppressesFormattingWhitespace(t *testing.T) {
	source := "<%\r\nResponse.Buffer = True\r\nResponse.BinaryWrite ChrB(&H42)\r\n%>\r\n \t\r\n<%\r\nResponse.BinaryWrite ChrB(&H4D)\r\nResponse.BinaryWrite ChrB(&H80)\r\n%>"
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

	expected := []byte{'B', 'M', 0x80}
	if !bytes.Equal(output.Bytes(), expected) {
		t.Fatalf("unexpected binary output with ASP block formatting: %v", output.Bytes())
	}
}

// TestVMResponseRedirectBareCall verifies that Response.Redirect terminates script execution
// even when inside a sub called without call keyword or parens (bare call).
func TestVMResponseRedirectBareCall(t *testing.T) {
	sourceBare := `<%
Response.Write "before"
RedirectMsg "index", "success"
Response.Write "after (should NOT appear)"

Sub RedirectMsg(msg, status)
    Response.Redirect msg & ".asp?status=" & status
End Sub
%>`

	compilerBare := NewASPCompiler(sourceBare)
	if err := compilerBare.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vmBare := NewVM(compilerBare.Bytecode(), compilerBare.Constants(), compilerBare.GlobalsCount())
	hostBare := NewMockHost()
	var outputBare bytes.Buffer
	hostBare.SetOutput(&outputBare)
	vmBare.SetHost(hostBare)

	if err := vmBare.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	hostBare.Response().Flush()

	if outputBare.String() != "" {
		t.Fatalf("unexpected output: %q (expected empty because Response.Redirect clears the buffer and terminates)", outputBare.String())
	}
	// Verify the redirect status was set.
	if status := hostBare.Response().GetStatus(); status != "302 Found" {
		t.Fatalf("expected 302 Found status after Response.Redirect, got %q", status)
	}
}

// TestVMResponseEndBareSubCall verifies that Response.End inside a bare Sub call (no Call keyword)
// terminates script execution immediately and no subsequent code runs.
func TestVMResponseEndBareSubCall(t *testing.T) {
	source := `<%
Dim afterFlag
afterFlag = False

Sub Inner()
    Response.End
End Sub

Inner
afterFlag = True
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	// Find the index of afterFlag
	afterFlagIdx, afterFlagExists := compiler.Globals.Get("afterFlag")
	if !afterFlagExists {
		t.Fatal("afterFlag global not found")
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}

	// If execution continued past the bare call, afterFlag would be True
	if vm.Globals[afterFlagIdx].Type == VTBool && vm.Globals[afterFlagIdx].Num != 0 {
		t.Fatal("BUG CONFIRMED: bare Sub call did NOT terminate execution — afterFlag was set to True after Response.End")
	}
}

// TestVMResponseRedirectBareSubCallWithOneArg verifies that a bare Sub call with 1 argument
// correctly propagates ResponseEndSignal.
func TestVMResponseRedirectBareSubCallWithOneArg(t *testing.T) {
	source := `<%
Dim afterFlag
afterFlag = False

Sub Inner(target)
    Response.Redirect target
End Sub

Inner "test.asp"
afterFlag = True
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	afterFlagIdx, afterFlagExists := compiler.Globals.Get("afterFlag")
	if !afterFlagExists {
		t.Fatal("afterFlag global not found")
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}

	// If execution continued past the bare call, afterFlag would be True
	if vm.Globals[afterFlagIdx].Type == VTBool && vm.Globals[afterFlagIdx].Num != 0 {
		t.Fatal("BUG CONFIRMED: bare Sub call with 1 arg did NOT terminate execution")
	}

	// Verify the redirect status was set.
	if status := host.Response().GetStatus(); status != "302 Found" {
		t.Fatalf("expected 302 Found status after Response.Redirect, got %q", status)
	}
}

// TestVMResponseRedirectBareSubCallWithTwoArgs verifies that a bare Sub call with 2 arguments
// correctly propagates ResponseEndSignal.
func TestVMResponseRedirectBareSubCallWithTwoArgs(t *testing.T) {
	source := `<%
Dim afterFlag
afterFlag = False

Sub Inner(target, msg)
    Response.Redirect target & ".asp?status=" & msg
End Sub

Inner "index", "success"
afterFlag = True
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	afterFlagIdx, afterFlagExists := compiler.Globals.Get("afterFlag")
	if !afterFlagExists {
		t.Fatal("afterFlag global not found")
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}

	// If execution continued past the bare call, afterFlag would be True
	if vm.Globals[afterFlagIdx].Type == VTBool && vm.Globals[afterFlagIdx].Num != 0 {
		t.Fatal("BUG CONFIRMED: bare Sub call with 2 args did NOT terminate execution")
	}

	// Verify the redirect status was set.
	if status := host.Response().GetStatus(); status != "302 Found" {
		t.Fatalf("expected 302 Found status after Response.Redirect, got %q", status)
	}
}

// TestVMResponseEndBareFunctionCall verifies that Response.End inside a bare Function call
// (no Call keyword, no parentheses) terminates script execution immediately.
func TestVMResponseEndBareFunctionCall(t *testing.T) {
	source := `<%
Dim afterFlag
afterFlag = False

Function Inner()
    Response.End
End Function

Inner
afterFlag = True
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	afterFlagIdx, afterFlagExists := compiler.Globals.Get("afterFlag")
	if !afterFlagExists {
		t.Fatal("afterFlag global not found")
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}

	// If execution continued past the bare call, afterFlag would be True
	if vm.Globals[afterFlagIdx].Type == VTBool && vm.Globals[afterFlagIdx].Num != 0 {
		t.Fatal("BUG CONFIRMED: bare Function call did NOT terminate execution")
	}
}

// TestVMResponseEndCallKeyword verifies that Call with parentheses works correctly (for comparison).
func TestVMResponseEndCallKeyword(t *testing.T) {
	source := `<%
Dim afterFlag
afterFlag = False

Sub Inner()
    Response.End
End Sub

Call Inner()
afterFlag = True
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	afterFlagIdx, afterFlagExists := compiler.Globals.Get("afterFlag")
	if !afterFlagExists {
		t.Fatal("afterFlag global not found")
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}

	// Call keyword should always work correctly
	if vm.Globals[afterFlagIdx].Type == VTBool && vm.Globals[afterFlagIdx].Num != 0 {
		t.Fatal("Call keyword also failed to terminate execution")
	}
}

// TestVMResponseEndNestedBareCalls verifies the exact bug report pattern: Outer (bare zero-arg call)
// calls InnerRedirect (bare call with 2 args) which calls Response.Redirect.
func TestVMResponseEndNestedBareCalls(t *testing.T) {
	source := `<%
Dim afterFlag
afterFlag = False

Sub InnerRedirect(target, msg)
    Response.Redirect target
End Sub

Sub Outer()
    InnerRedirect "target.asp", "success"
    afterFlag = True
End Sub

Outer
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	afterFlagIdx, afterFlagExists := compiler.Globals.Get("afterFlag")
	if !afterFlagExists {
		t.Fatal("afterFlag global not found")
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}

	// afterFlag is set inside Outer AFTER the InnerRedirect call.
	// If ResponseEndSignal propagated correctly, afterFlag should remain False.
	if vm.Globals[afterFlagIdx].Type == VTBool && vm.Globals[afterFlagIdx].Num != 0 {
		t.Fatal("BUG CONFIRMED: nested bare calls did NOT terminate execution — afterFlag was set to True")
	}

	// Verify the redirect status.
	if status := host.Response().GetStatus(); status != "302 Found" {
		t.Fatalf("expected 302 Found status after Response.Redirect, got %q", status)
	}
}

// TestVMResponseEndBareCallForwardRef tests a bare Sub call where the Sub is defined after the call site.
func TestVMResponseEndBareCallForwardRef(t *testing.T) {
	source := `<%
Dim afterFlag
afterFlag = False

Inner "test.asp"
afterFlag = True

Sub Inner(target)
    Response.Redirect target
End Sub
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	afterFlagIdx, afterFlagExists := compiler.Globals.Get("afterFlag")
	if !afterFlagExists {
		t.Fatal("afterFlag global not found")
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}

	if vm.Globals[afterFlagIdx].Type == VTBool && vm.Globals[afterFlagIdx].Num != 0 {
		t.Fatal("BUG CONFIRMED: forward reference bare call did NOT terminate execution")
	}
}

// TestVMResponseEndBareCallInsideIfBlock tests a bare Sub call with args inside an If block.
func TestVMResponseEndBareCallInsideIfBlock(t *testing.T) {
	source := `<%
Dim afterFlag
afterFlag = False

Sub Inner(target)
    Response.Redirect target
End Sub

If True Then
    Inner "test.asp"
    afterFlag = True
End If
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	afterFlagIdx, afterFlagExists := compiler.Globals.Get("afterFlag")
	if !afterFlagExists {
		t.Fatal("afterFlag global not found")
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}

	if vm.Globals[afterFlagIdx].Type == VTBool && vm.Globals[afterFlagIdx].Num != 0 {
		t.Fatal("BUG CONFIRMED: bare call inside If block did NOT terminate execution")
	}

	if status := host.Response().GetStatus(); status != "302 Found" {
		t.Fatalf("expected 302 Found status after Response.Redirect, got %q", status)
	}
}

// TestVMResponseRedirectBareMemberCall tests Response.Redirect as a bare member call (no parens).
func TestVMResponseRedirectBareMemberCall(t *testing.T) {
	source := `<%
Dim afterFlag
afterFlag = False
Response.Redirect "test.asp"
afterFlag = True
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	afterFlagIdx, afterFlagExists := compiler.Globals.Get("afterFlag")
	if !afterFlagExists {
		t.Fatal("afterFlag global not found")
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}

	if vm.Globals[afterFlagIdx].Type == VTBool && vm.Globals[afterFlagIdx].Num != 0 {
		t.Fatal("BUG CONFIRMED: bare member call Response.Redirect did NOT terminate execution")
	}

	if status := host.Response().GetStatus(); status != "302 Found" {
		t.Fatalf("expected 302 Found status after Response.Redirect, got %q", status)
	}
}

// TestVMResponseEndBareMemberCall tests Response.End as a bare member call (no parens).
func TestVMResponseEndBareMemberCall(t *testing.T) {
	source := `<%
Dim afterFlag
afterFlag = False
Response.End
afterFlag = True
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	afterFlagIdx, afterFlagExists := compiler.Globals.Get("afterFlag")
	if !afterFlagExists {
		t.Fatal("afterFlag global not found")
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}

	if vm.Globals[afterFlagIdx].Type == VTBool && vm.Globals[afterFlagIdx].Num != 0 {
		t.Fatal("BUG CONFIRMED: bare member call Response.End did NOT terminate execution")
	}
}

// TestVMResponseEndDeeplyNestedBareCalls tests 3 levels of nested bare calls.
func TestVMResponseEndDeeplyNestedBareCalls(t *testing.T) {
	source := `<%
Dim afterFlag
afterFlag = False

Sub Level3()
    Response.End
End Sub

Sub Level2()
    Level3
End Sub

Sub Level1()
    Level2
End Sub

Level1
afterFlag = True
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	afterFlagIdx, afterFlagExists := compiler.Globals.Get("afterFlag")
	if !afterFlagExists {
		t.Fatal("afterFlag global not found")
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}

	if vm.Globals[afterFlagIdx].Type == VTBool && vm.Globals[afterFlagIdx].Num != 0 {
		t.Fatal("BUG CONFIRMED: deeply nested bare calls did NOT terminate execution")
	}
}

// TestVMResponseEndBareCallInLoop tests a bare Sub call with args inside a loop.
func TestVMResponseEndBareCallInLoop(t *testing.T) {
	source := `<%
Dim afterFlag
afterFlag = False

Sub Inner()
    Response.End
End Sub

Dim i
For i = 1 To 3
    Inner
    afterFlag = True
Next
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	afterFlagIdx, afterFlagExists := compiler.Globals.Get("afterFlag")
	if !afterFlagExists {
		t.Fatal("afterFlag global not found")
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}

	if vm.Globals[afterFlagIdx].Type == VTBool && vm.Globals[afterFlagIdx].Num != 0 {
		t.Fatal("BUG CONFIRMED: bare call in loop did NOT terminate execution")
	}
}

// TestVMResponseEndBareCallOnErrorResumeNext tests bare call termination with On Error Resume Next active.
func TestVMResponseEndBareCallOnErrorResumeNext(t *testing.T) {
	source := `<%
Dim afterFlag
afterFlag = False

Sub Inner()
    Response.End
End Sub

On Error Resume Next
Inner
afterFlag = True
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	afterFlagIdx, afterFlagExists := compiler.Globals.Get("afterFlag")
	if !afterFlagExists {
		t.Fatal("afterFlag global not found")
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}

	// Response.End should still terminate even with On Error Resume Next
	if vm.Globals[afterFlagIdx].Type == VTBool && vm.Globals[afterFlagIdx].Num != 0 {
		t.Fatal("BUG CONFIRMED: bare call with On Error Resume Next did NOT terminate execution")
	}
}

// TestVMResponseRedirectMixedCallAndBare tests a top-level Call with parens that internally
// calls a Sub via bare call that does Response.Redirect.
func TestVMResponseRedirectMixedCallAndBare(t *testing.T) {
	source := `<%
Dim afterFlag
afterFlag = False

Sub Inner(target)
    Response.Redirect target
End Sub

Sub Outer()
    Inner "test.asp"
    afterFlag = True
End Sub

Call Outer()
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	afterFlagIdx, afterFlagExists := compiler.Globals.Get("afterFlag")
	if !afterFlagExists {
		t.Fatal("afterFlag global not found")
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}

	if vm.Globals[afterFlagIdx].Type == VTBool && vm.Globals[afterFlagIdx].Num != 0 {
		t.Fatal("BUG CONFIRMED: mixed Call/bare did NOT terminate execution")
	}

	if status := host.Response().GetStatus(); status != "302 Found" {
		t.Fatalf("expected 302 Found status after Response.Redirect, got %q", status)
	}
}

// TestVMResponseEndBareSubCallMultiBlock verifies that a bare Sub call
// terminates script execution even when the Sub declaration and the call site are in
// separate <% %> script blocks.
func TestVMResponseEndBareSubCallMultiBlock(t *testing.T) {
	source := `<%
Dim afterFlag
afterFlag = False

Sub Inner()
    Response.End
End Sub
%><%
Inner
afterFlag = True
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	afterFlagIdx, afterFlagExists := compiler.Globals.Get("afterFlag")
	if !afterFlagExists {
		t.Fatal("afterFlag global not found")
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}

	if vm.Globals[afterFlagIdx].Type == VTBool && vm.Globals[afterFlagIdx].Num != 0 {
		t.Fatal("BUG: bare Sub call across block boundary did NOT terminate execution — afterFlag was set to True")
	}
}

// TestVMResponseEndBareSubCallMultiBlockWithArgs verifies that a bare Sub call WITH arguments
// terminates script execution across <% %> block boundaries.
func TestVMResponseEndBareSubCallMultiBlockWithArgs(t *testing.T) {
	source := `<%
Dim afterFlag
afterFlag = False

Sub Inner(x)
    Response.End
End Sub
%><%
Inner "test"
afterFlag = True
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	afterFlagIdx, afterFlagExists := compiler.Globals.Get("afterFlag")
	if !afterFlagExists {
		t.Fatal("afterFlag global not found")
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}

	if vm.Globals[afterFlagIdx].Type == VTBool && vm.Globals[afterFlagIdx].Num != 0 {
		t.Fatal("BUG: bare Sub call with args across block boundary did NOT terminate execution — afterFlag was set to True")
	}
}

// TestVMResponseRedirectBareSubCallTwoArgsMultiBlock verifies that Response.Redirect inside a bare Sub call
// with 2 arguments terminates script execution across <% %> block boundaries.
func TestVMResponseRedirectBareSubCallTwoArgsMultiBlock(t *testing.T) {
	source := `<%
Dim afterFlag
afterFlag = False

Sub Inner(target, msg)
    Response.Redirect target & ".asp?status=" & msg
End Sub
%><%
Inner "index", "success"
afterFlag = True
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	afterFlagIdx, afterFlagExists := compiler.Globals.Get("afterFlag")
	if !afterFlagExists {
		t.Fatal("afterFlag global not found")
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}

	if vm.Globals[afterFlagIdx].Type == VTBool && vm.Globals[afterFlagIdx].Num != 0 {
		t.Fatal("BUG: bare Sub call with 2 args across block boundary did NOT terminate execution — afterFlag was set to True")
	}

	if status := host.Response().GetStatus(); status != "302 Found" {
		t.Fatalf("expected 302 Found status after Response.Redirect, got %q", status)
	}
}

// TestVMResponseEndBareSubCallForwardRefMultiBlock verifies a forward-reference bare Sub call
// (Sub declared after the call site) across <% %> block boundaries still terminates.
func TestVMResponseEndBareSubCallForwardRefMultiBlock(t *testing.T) {
	source := `<%
Dim afterFlag
afterFlag = False
%><%
Inner "test"
afterFlag = True
%><%
Sub Inner(target)
    Response.End
End Sub
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	afterFlagIdx, afterFlagExists := compiler.Globals.Get("afterFlag")
	if !afterFlagExists {
		t.Fatal("afterFlag global not found")
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}

	if vm.Globals[afterFlagIdx].Type == VTBool && vm.Globals[afterFlagIdx].Num != 0 {
		t.Fatal("BUG: forward-ref bare Sub call with args across block boundary did NOT terminate execution — afterFlag was set to True")
	}
}

// TestVMResponseRedirectBareSubCallTwoArgsMultiBlockWithGap verifies that Response.Redirect
// inside a bare Sub call with 2 arguments terminates execution across <% %> blocks that
// have HTML content/gap between them (simulating #include boundary).
func TestVMResponseRedirectBareSubCallTwoArgsMultiBlockWithGap(t *testing.T) {
	source := `<%
Sub Inner(target, msg)
    Response.Redirect "target.asp?act=" & target & "&msg=" & msg
End Sub
%>
<%
Response.Write "before<br>"
Inner "index", "success"
Response.Write "after (should NOT appear)<br>"
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

	// With Response.Redirect, the buffer is reset so no output should appear.
	if output.String() != "" {
		t.Fatalf("unexpected output: %q (expected empty because Response.Redirect clears the buffer and terminates)", output.String())
	}

	if status := host.Response().GetStatus(); status != "302 Found" {
		t.Fatalf("expected 302 Found status after Response.Redirect, got %q", status)
	}
}

// TestVMResponseRedirectBareSubCallTwoArgsMultiBlockWithGapForwardRef is like
// TestVMResponseRedirectBareSubCallTwoArgsMultiBlockWithGap but the Sub is
// declared AFTER the call site (forward reference).
func TestVMResponseRedirectBareSubCallTwoArgsMultiBlockWithGapForwardRef(t *testing.T) {
	source := `<%
Response.Write "before<br>"
%>
<%
Inner "index", "success"
Response.Write "after (should NOT appear)<br>"
%>
<%
Sub Inner(target, msg)
    Response.Redirect "target.asp?act=" & target & "&msg=" & msg
End Sub
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

	// With Response.Redirect, the buffer is reset so no output should appear.
	if output.String() != "" {
		t.Fatalf("unexpected output: %q (expected empty because Response.Redirect clears the buffer and terminates)", output.String())
	}

	if status := host.Response().GetStatus(); status != "302 Found" {
		t.Fatalf("expected 302 Found status after Response.Redirect, got %q", status)
	}
}

// TestVMResponseRedirectBareSubCallTwoArgsMultiBlockHTMLWrite verifies Response.Redirect
// across blocks where the second block uses <%= expression output between the sub definition
// and the bare call.
func TestVMResponseRedirectBareSubCallTwoArgsMultiBlockHTMLWrite(t *testing.T) {
	source := `<%
Sub Inner(target, msg)
    Response.Redirect "target.asp?act=" & target & "&msg=" & msg
End Sub
%>
<%= "middle<br>" %>
<%
Inner "index", "success"
Response.Write "after (should NOT appear)<br>"
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

	// Response.Redirect clears the buffer, so all output before it is discarded.
	if output.String() != "" {
		t.Fatalf("unexpected output: %q (expected empty because Response.Redirect clears the buffer)", output.String())
	}

	if status := host.Response().GetStatus(); status != "302 Found" {
		t.Fatalf("expected 302 Found status after Response.Redirect, got %q", status)
	}
}
