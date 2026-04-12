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
	"bytes"
	"os"
	"path/filepath"
	"strings"
	"testing"
)

// TestVMServerCreateObjectG3Files verifies native object creation for G3FILES.
func TestVMServerCreateObjectG3Files(t *testing.T) {
	vm := NewVM(nil, nil, 5)
	host := NewMockHost()
	host.Server().SetRootDir(t.TempDir())
	host.Server().SetRequestPath("/tests/test_g3files.asp")
	vm.SetHost(host)

	obj := vm.dispatchNativeCall(nativeObjectServer, "CreateObject", []Value{NewString("G3FILES")})
	if obj.Type != VTNativeObject {
		t.Fatalf("expected VTNativeObject, got %#v", obj)
	}
}

// TestG3FilesEncodingAndEOL verifies write/read with UTF-16 BOM and line ending normalization.
func TestG3FilesEncodingAndEOL(t *testing.T) {
	vm := NewVM(nil, nil, 5)
	host := NewMockHost()
	rootDir := t.TempDir()
	host.Server().SetRootDir(rootDir)
	host.Server().SetRequestPath("/tests/test_g3files.asp")
	vm.SetHost(host)

	obj := vm.dispatchNativeCall(nativeObjectServer, "CreateObject", []Value{NewString("G3FILES")})
	ok := vm.dispatchNativeCall(obj.Num, "Write", []Value{
		NewString("/sandbox/utf16.txt"),
		NewString("line1\nline2"),
		NewString("utf-16le"),
		NewString("windows"),
		NewBool(true),
	})
	if ok.Type != VTBool || ok.Num != 1 {
		t.Fatalf("expected Write=True, got %#v", ok)
	}

	absPath := filepath.Join(rootDir, "sandbox", "utf16.txt")
	raw, err := os.ReadFile(absPath)
	if err != nil {
		t.Fatalf("read raw file: %v", err)
	}
	if len(raw) < 2 || !bytes.Equal(raw[:2], []byte{0xFF, 0xFE}) {
		t.Fatalf("expected UTF-16LE BOM, got %v", raw)
	}

	read := vm.dispatchNativeCall(obj.Num, "Read", []Value{NewString("/sandbox/utf16.txt"), NewString("utf-16le")})
	if read.Type != VTString {
		t.Fatalf("expected VTString, got %#v", read)
	}
	if read.Str != "line1\r\nline2" {
		t.Fatalf("unexpected decoded content: %q", read.Str)
	}
}

// TestG3FilesConvertEncoding verifies file conversion between UTF-16 and UTF-8.
func TestG3FilesConvertEncoding(t *testing.T) {
	vm := NewVM(nil, nil, 5)
	host := NewMockHost()
	rootDir := t.TempDir()
	host.Server().SetRootDir(rootDir)
	host.Server().SetRequestPath("/tests/test_g3files.asp")
	vm.SetHost(host)

	obj := vm.dispatchNativeCall(nativeObjectServer, "CreateObject", []Value{NewString("G3FILES")})
	_ = vm.dispatchNativeCall(obj.Num, "Write", []Value{
		NewString("/sandbox/source.txt"),
		NewString("áéí"),
		NewString("utf-16le"),
		NewString("linux"),
		NewBool(true),
	})

	converted := vm.dispatchNativeCall(obj.Num, "ConvertFileEncoding", []Value{
		NewString("/sandbox/source.txt"),
		NewString("/sandbox/target.txt"),
		NewString("utf-16le"),
		NewString("utf-8"),
		NewString("linux"),
		NewBool(false),
	})
	if converted.Type != VTBool || converted.Num != 1 {
		t.Fatalf("expected ConvertFileEncoding=True, got %#v", converted)
	}

	read := vm.dispatchNativeCall(obj.Num, "Read", []Value{NewString("/sandbox/target.txt"), NewString("utf-8")})
	if read.Type != VTString || !strings.Contains(read.Str, "áéí") {
		t.Fatalf("unexpected converted file content: %#v", read)
	}
}
