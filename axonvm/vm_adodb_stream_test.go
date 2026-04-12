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

// TestVMServerCreateObjectADODBStream verifies Server.CreateObject native ADODB.Stream integration.
func TestVMServerCreateObjectADODBStream(t *testing.T) {
	vm := NewVM(nil, nil, 5)
	host := NewMockHost()
	host.Server().SetRootDir(t.TempDir())
	host.Server().SetRequestPath("/tests/test_adodb.asp")
	vm.SetHost(host)

	stream := vm.dispatchNativeCall(nativeObjectServer, "CreateObject", []Value{NewString("ADODB.Stream")})
	if stream.Type != VTNativeObject {
		t.Fatalf("expected VTNativeObject, got %#v", stream)
	}

	vm.dispatchMemberSet(stream.Num, "Type", NewInteger(2))
	vm.dispatchMemberSet(stream.Num, "Charset", NewString("utf-8"))
	vm.dispatchNativeCall(stream.Num, "Open", nil)
	vm.dispatchNativeCall(stream.Num, "WriteText", []Value{NewString("AxonASP")})
	vm.dispatchMemberSet(stream.Num, "Position", NewInteger(0))

	text := vm.dispatchMemberGet(stream, "ReadText")
	if text.Type != VTString || text.Str != "AxonASP" {
		t.Fatalf("unexpected ReadText content: %#v", text)
	}

	eos := vm.dispatchMemberGet(stream, "EOS")
	if eos.Type != VTBool || eos.Num != 1 {
		t.Fatalf("unexpected EOS state: %#v", eos)
	}
}

// TestVMADODBStreamSaveAndLoad verifies SaveToFile and LoadFromFile with sandboxed paths.
func TestVMADODBStreamSaveAndLoad(t *testing.T) {
	vm := NewVM(nil, nil, 5)
	host := NewMockHost()
	rootDir := t.TempDir()
	host.Server().SetRootDir(rootDir)
	host.Server().SetRequestPath("/tests/test_adodb.asp")
	vm.SetHost(host)

	if err := os.MkdirAll(filepath.Join(rootDir, "sandbox"), 0755); err != nil {
		t.Fatalf("failed to create sandbox directory: %v", err)
	}

	writer := vm.dispatchNativeCall(nativeObjectServer, "CreateObject", []Value{NewString("ADODB.Stream")})
	vm.dispatchNativeCall(writer.Num, "Open", nil)
	vm.dispatchMemberSet(writer.Num, "Charset", NewString("utf-8"))
	vm.dispatchNativeCall(writer.Num, "WriteText", []Value{NewString("File Content")})
	vm.dispatchNativeCall(writer.Num, "SaveToFile", []Value{NewString("/sandbox/sample.txt"), NewInteger(2)})

	reader := vm.dispatchNativeCall(nativeObjectServer, "CreateObject", []Value{NewString("ADODB.Stream")})
	vm.dispatchNativeCall(reader.Num, "Open", nil)
	vm.dispatchMemberSet(reader.Num, "Charset", NewString("utf-8"))
	vm.dispatchNativeCall(reader.Num, "LoadFromFile", []Value{NewString("/sandbox/sample.txt")})

	value := vm.dispatchNativeCall(reader.Num, "ReadText", nil)
	if value.Type != VTString || value.Str != "File Content" {
		t.Fatalf("unexpected loaded stream content: %#v", value)
	}
}

// TestVMADODBStreamCopyToAndSetEOS verifies CopyTo transfer and SetEOS truncation behavior.
func TestVMADODBStreamCopyToAndSetEOS(t *testing.T) {
	vm := NewVM(nil, nil, 5)
	host := NewMockHost()
	host.Server().SetRootDir(t.TempDir())
	host.Server().SetRequestPath("/tests/test_adodb.asp")
	vm.SetHost(host)

	source := vm.dispatchNativeCall(nativeObjectServer, "CreateObject", []Value{NewString("ADODB.Stream")})
	destination := vm.dispatchNativeCall(nativeObjectServer, "CreateObject", []Value{NewString("ADODB.Stream")})

	vm.dispatchNativeCall(source.Num, "Open", nil)
	vm.dispatchNativeCall(destination.Num, "Open", nil)
	vm.dispatchMemberSet(source.Num, "Charset", NewString("utf-8"))
	vm.dispatchMemberSet(destination.Num, "Charset", NewString("utf-8"))

	vm.dispatchNativeCall(source.Num, "WriteText", []Value{NewString("ABCDEF")})
	vm.dispatchMemberSet(source.Num, "Position", NewInteger(0))

	vm.dispatchNativeCall(source.Num, "CopyTo", []Value{destination, NewInteger(3)})
	vm.dispatchMemberSet(destination.Num, "Position", NewInteger(0))

	partial := vm.dispatchNativeCall(destination.Num, "ReadText", nil)
	if partial.Type != VTString || partial.Str != "ABC" {
		t.Fatalf("unexpected CopyTo result: %#v", partial)
	}

	vm.dispatchMemberSet(source.Num, "Position", NewInteger(3))
	vm.dispatchNativeCall(source.Num, "SetEOS", nil)

	size := vm.dispatchMemberGet(source, "Size")
	if size.Type != VTInteger || size.Num != 3 {
		t.Fatalf("unexpected Size after SetEOS: %#v", size)
	}
}

// TestASPADODBStreamReadFilePattern verifies the same ADODB.Stream flow used by manual/default.asp.
func TestASPADODBStreamReadFilePattern(t *testing.T) {
	rootDir := t.TempDir()

	source := `<%
Function ReadFile(path)
	Dim stream
	Set stream = Server.CreateObject("ADODB.Stream")
	stream.Type = 2
	stream.Charset = "utf-8"
	stream.Open
	stream.WriteText "hello-manual"
	stream.Position = 0
	ReadFile = stream.ReadText
	stream.Close
End Function

Response.Write ReadFile("unused")
%>`

	compiler := NewASPCompiler(source)
	compiler.SetSourceName("www/manual/default.asp")
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	host.Server().SetRootDir(rootDir)
	host.Server().SetRequestPath("/manual/default.asp")
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}

	_ = output.String()
}
