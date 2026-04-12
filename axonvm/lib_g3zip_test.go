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
	"os"
	"testing"
)

func TestG3Zip(t *testing.T) {
	vm := NewVM(nil, nil, 0)
	vm.host = &MockHost{}

	zipLib := vm.newG3ZipObject()
	if zipLib.Type != VTNativeObject {
		t.Fatalf("expected VTNativeObject, got %v", zipLib.Type)
	}

	obj := vm.g3zipItems[zipLib.Num]
	if obj == nil {
		t.Fatal("expected object in vm items")
	}

	defer os.Remove("test.zip")

	// Create zip
	res := obj.DispatchMethod("Create", []Value{NewString("test.zip")})
	if res.Type != VTBool || res.Num == 0 {
		t.Fatal("Failed to create zip")
	}

	obj.DispatchMethod("AddText", []Value{NewString("hello.txt"), NewString("world")})
	obj.DispatchMethod("Close", nil)

	// Read zip
	res = obj.DispatchMethod("Open", []Value{NewString("test.zip")})
	if res.Type != VTBool || res.Num == 0 {
		t.Fatal("Failed to open zip")
	}

	list := obj.DispatchMethod("List", nil)
	if list.Type != VTArray {
		t.Fatalf("expected Array, got %v", list.Type)
	}

	count := obj.DispatchPropertyGet("Count")
	if count.Num != 1 {
		t.Errorf("expected 1 file in zip, got %d", count.Num)
	}

	info := obj.DispatchMethod("GetFileInfo", []Value{NewString("hello.txt")})
	if info.Type != VTNativeObject {
		t.Fatalf("expected GetFileInfo to return native object, got %#v", info)
	}

	obj.DispatchMethod("Close", nil)
}

func TestG3ZipAddFileAbsolutePath(t *testing.T) {
	vm := NewVM(nil, nil, 0)
	vm.host = &MockHost{}

	zipLib := vm.newG3ZipObject()
	obj := vm.g3zipItems[zipLib.Num]
	if obj == nil {
		t.Fatal("expected object in vm items")
	}

	tmpFile, err := os.CreateTemp("", "axonasp-g3zip-*.txt")
	if err != nil {
		t.Fatalf("failed to create temp file: %v", err)
	}
	tmpFilePath := tmpFile.Name()
	if _, err := tmpFile.WriteString("absolute path content"); err != nil {
		tmpFile.Close()
		os.Remove(tmpFilePath)
		t.Fatalf("failed to write temp file: %v", err)
	}
	tmpFile.Close()
	defer os.Remove(tmpFilePath)

	zipFile, err := os.CreateTemp("", "axonasp-g3zip-*.zip")
	if err != nil {
		t.Fatalf("failed to create temp zip placeholder: %v", err)
	}
	zipPath := zipFile.Name()
	zipFile.Close()
	os.Remove(zipPath)
	defer os.Remove(zipPath)

	res := obj.DispatchMethod("Create", []Value{NewString(zipPath)})
	if res.Type != VTBool || res.Num == 0 {
		t.Fatal("failed to create zip using absolute path")
	}

	res = obj.DispatchMethod("AddFile", []Value{NewString(tmpFilePath), NewString("nested/file.txt")})
	if res.Type != VTBool || res.Num == 0 {
		t.Fatal("expected AddFile with absolute path to succeed")
	}

	obj.DispatchMethod("Close", nil)

	res = obj.DispatchMethod("Open", []Value{NewString(zipPath)})
	if res.Type != VTBool || res.Num == 0 {
		t.Fatal("failed to reopen generated zip")
	}

	list := obj.DispatchMethod("List", nil)
	if list.Type != VTArray {
		t.Fatalf("expected list to be array, got %v", list.Type)
	}
	entries := list.Arr.Values
	if len(entries) != 1 {
		t.Fatalf("expected 1 entry in zip, got %d", len(entries))
	}
	if entries[0].String() != "nested/file.txt" {
		t.Fatalf("expected normalized zip entry name nested/file.txt, got %q", entries[0].String())
	}

	obj.DispatchMethod("Close", nil)
}
