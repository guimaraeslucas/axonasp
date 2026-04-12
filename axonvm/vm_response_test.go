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
	"os"
	"path/filepath"
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

// TestVMCaptchaPageProducesBitmapSignature verifies the legacy CAPTCHA page emits a valid BMP signature and binary response type.
func TestVMCaptchaPageProducesBitmapSignature(t *testing.T) {
	sourcePath := filepath.Join("..", "www", "tests", "captcha.asp")
	absPath, err := filepath.Abs(sourcePath)
	if err != nil {
		t.Fatalf("abs path failed: %v", err)
	}
	sourceBytes, err := os.ReadFile(absPath)
	if err != nil {
		t.Fatalf("read captcha source failed: %v", err)
	}

	compiler := NewASPCompiler(string(sourceBytes))
	compiler.SetSourceName(absPath)
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

	if host.Response().GetContentType() != "image/bmp" {
		t.Fatalf("expected image/bmp content type, got %q", host.Response().GetContentType())
	}
	if len(output.Bytes()) < 2 || !bytes.Equal(output.Bytes()[:2], []byte{'B', 'M'}) {
		t.Fatalf("expected BMP signature, got %v", output.Bytes())
	}
}
