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
	"fmt"
	"os"
	"path/filepath"
	"testing"
	"time"
)

func TestG3TARCreateListExtract(t *testing.T) {
	vm := NewVM(nil, nil, 0)
	vm.host = NewMockHost()

	objVal := vm.newG3TARObject()
	obj := vm.g3tarItems[objVal.Num]
	if obj == nil {
		t.Fatal("expected g3tar native object")
	}

	relBase := fmt.Sprintf("temp/test-g3tar-%d", time.Now().UnixNano())
	archiveRel := filepath.ToSlash(filepath.Join(relBase, "sample.tar"))
	srcRel := filepath.ToSlash(filepath.Join(relBase, "input.txt"))
	srcPath, ok := vm.fsoResolvePath(srcRel)
	if !ok {
		t.Fatal("failed to resolve source path in sandbox")
	}
	if err := os.MkdirAll(filepath.Dir(srcPath), 0755); err != nil {
		t.Fatalf("failed to create test directory: %v", err)
	}
	if err := os.WriteFile(srcPath, []byte("tar-content"), 0644); err != nil {
		t.Fatalf("failed to prepare input file: %v", err)
	}
	defer os.RemoveAll(filepath.Dir(srcPath))

	if res := obj.DispatchMethod("Create", []Value{NewString(archiveRel)}); res.Type != VTBool || res.Num == 0 {
		t.Fatal("expected Create to succeed")
	}
	if res := obj.DispatchMethod("AddFile", []Value{NewString(srcRel), NewString("folder/input.txt")}); res.Type != VTBool || res.Num == 0 {
		t.Fatal("expected AddFile to succeed")
	}
	if res := obj.DispatchMethod("AddText", []Value{NewString("readme.txt"), NewString("hello")}); res.Type != VTBool || res.Num == 0 {
		t.Fatal("expected AddText to succeed")
	}
	obj.DispatchMethod("Close", nil)

	if res := obj.DispatchMethod("Open", []Value{NewString(archiveRel)}); res.Type != VTBool || res.Num == 0 {
		t.Fatal("expected Open to succeed")
	}
	list := obj.DispatchMethod("List", nil)
	if list.Type != VTArray || list.Arr == nil {
		t.Fatalf("expected list array, got %#v", list)
	}
	if len(list.Arr.Values) != 2 {
		t.Fatalf("expected 2 tar entries, got %d", len(list.Arr.Values))
	}

	extractRel := filepath.ToSlash(filepath.Join(relBase, "out"))
	if res := obj.DispatchMethod("ExtractAll", []Value{NewString(extractRel)}); res.Type != VTBool || res.Num == 0 {
		t.Fatal("expected ExtractAll to succeed")
	}
	extractPath, ok := vm.fsoResolvePath(extractRel)
	if !ok {
		t.Fatal("failed to resolve extraction path")
	}

	payload, err := os.ReadFile(filepath.Join(extractPath, "folder", "input.txt"))
	if err != nil {
		t.Fatalf("failed to read extracted file: %v", err)
	}
	if string(payload) != "tar-content" {
		t.Fatalf("unexpected extracted payload: %q", string(payload))
	}
}
