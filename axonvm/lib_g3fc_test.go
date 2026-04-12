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
	"os"
	"path/filepath"
	"testing"
)

func TestG3FC_CreateAndExtract(t *testing.T) {
	vm := NewVM(nil, nil, 5)
	host := NewMockHost()
	rootDir := t.TempDir()
	host.Server().SetRootDir(rootDir)
	host.Server().SetRequestPath("/tests/test_g3fc.asp")
	vm.SetHost(host)

	// Create a dummy file to archive
	sandboxDir := filepath.Join(rootDir, "sandbox")
	os.MkdirAll(sandboxDir, 0755)
	testFile := filepath.Join(sandboxDir, "test.txt")
	os.WriteFile(testFile, []byte("G3FC test content"), 0644)

	g3fc := vm.newG3FCObject()

	// 1. Test Create
	archivePath := "/sandbox/test.g3fc"
	createArgs := []Value{
		NewString(archivePath),
		NewString("/sandbox/test.txt"),
	}
	res := vm.dispatchNativeCall(g3fc.Num, "Create", createArgs)
	if res.Type != VTBool || res.Num == 0 {
		t.Fatalf("expected Create to return True, got %#v", res)
	}

	// 2. Test List
	listArgs := []Value{NewString(archivePath)}
	listRes := vm.dispatchNativeCall(g3fc.Num, "List", listArgs)
	if listRes.Type != VTArray || listRes.Arr == nil || len(listRes.Arr.Values) == 0 {
		t.Fatalf("expected List to return array with entries, got %#v", listRes)
	}

	found := false
	for _, v := range listRes.Arr.Values {
		if v.Type == VTNativeObject {
			pathVal, ok := vm.dispatchDictionaryMethod(v.Num, "Item", []Value{NewString("Path")})
			if ok && pathVal.Str == "test.txt" {
				found = true
				break
			}
		}
	}
	if !found {
		t.Errorf("test.txt not found in archive listing")
	}

	// 3. Test Extract
	extractDir := "/sandbox/extracted"
	extractArgs := []Value{
		NewString(archivePath),
		NewString(extractDir),
	}
	res = vm.dispatchNativeCall(g3fc.Num, "Extract", extractArgs)
	if res.Type != VTBool || res.Num == 0 {
		t.Fatalf("expected Extract to return True, got %#v", res)
	}

	// Verify extracted file
	extractedFile := filepath.Join(sandboxDir, "extracted", "test.txt")
	content, err := os.ReadFile(extractedFile)
	if err != nil {
		t.Fatalf("failed to read extracted file: %v", err)
	}
	if string(content) != "G3FC test content" {
		t.Errorf("expected 'G3FC test content', got %q", string(content))
	}
}

func TestG3FC_Encryption(t *testing.T) {
	vm := NewVM(nil, nil, 5)
	host := NewMockHost()
	rootDir := t.TempDir()
	host.Server().SetRootDir(rootDir)
	host.Server().SetRequestPath("/tests/test_g3fc_enc.asp")
	vm.SetHost(host)

	sandboxDir := filepath.Join(rootDir, "sandbox")
	os.MkdirAll(sandboxDir, 0755)
	testFile := filepath.Join(sandboxDir, "secret.txt")
	os.WriteFile(testFile, []byte("Sensitive Data"), 0644)

	g3fc := vm.newG3FCObject()
	archivePath := "/sandbox/secret.g3fc"
	password := "AxonPass"

	// 1. Create with Password
	createArgs := []Value{
		NewString(archivePath),
		NewString("/sandbox/secret.txt"),
		NewString(password),
	}
	res := vm.dispatchNativeCall(g3fc.Num, "Create", createArgs)
	if res.Type != VTBool || res.Num == 0 {
		t.Fatalf("Create failed")
	}

	// 2. Extract with WRONG password should fail or raise error
	// Set onResumeNext to true to prevent panic on raised error
	vm.onResumeNext = true
	extractArgsWrong := []Value{
		NewString(archivePath),
		NewString("/sandbox/wrong"),
		NewString("WrongPass"),
	}
	res = vm.dispatchNativeCall(g3fc.Num, "Extract", extractArgsWrong)

	if res.Type == VTBool && res.Num == 1 {
		t.Errorf("expected Extract with wrong password to fail")
	}
	if vm.lastError == nil {
		t.Errorf("expected lastError to be set on decryption failure")
	}
	vm.onResumeNext = false
	vm.lastError = nil

	// 3. Extract with CORRECT password
	extractArgsCorrect := []Value{
		NewString(archivePath),
		NewString("/sandbox/correct"),
		NewString(password),
	}
	res = vm.dispatchNativeCall(g3fc.Num, "Extract", extractArgsCorrect)
	if res.Type != VTBool || res.Num == 0 {
		t.Fatalf("Extract with correct password failed: %v", vm.lastError)
	}

	extractedFile := filepath.Join(sandboxDir, "correct", "secret.txt")
	content, _ := os.ReadFile(extractedFile)
	if string(content) != "Sensitive Data" {
		t.Errorf("expected 'Sensitive Data', got %q", string(content))
	}
}
