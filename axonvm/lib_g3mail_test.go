//go:build !wasm

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

func TestG3Mail(t *testing.T) {
	vm := NewVM(nil, nil, 0)

	mailLib := vm.newG3MailObject()
	if mailLib.Type != VTNativeObject {
		t.Fatalf("expected VTNativeObject, got %v", mailLib.Type)
	}

	obj := vm.g3mailItems[mailLib.Num]

	// Test properties
	obj.DispatchPropertySet("Subject", []Value{NewString("Hello")})
	subj := obj.DispatchPropertyGet("Subject")
	if subj.String() != "Hello" {
		t.Errorf("expected Hello, got %s", subj.String())
	}

	obj.DispatchPropertySet("IsHTML", []Value{NewBool(true)})
	isHtml := obj.DispatchPropertyGet("IsHTML")
	if isHtml.Type != VTBool || isHtml.Num == 0 {
		t.Error("expected IsHTML to be true")
	}

	// Test HTMLBody property get/set
	obj.DispatchPropertySet("HTMLBody", []Value{NewString("<h1>Hello HTML</h1>")})
	htmlBody := obj.DispatchPropertyGet("HTMLBody")
	if htmlBody.String() != "<h1>Hello HTML</h1>" {
		t.Errorf("expected HTMLBody to be '<h1>Hello HTML</h1>', got %s", htmlBody.String())
	}

	obj.DispatchMethod("AddAddress", []Value{NewString("test@example.com")})
	if len(obj.to) != 1 || obj.to[0] != "test@example.com" {
		t.Error("AddAddress failed")
	}

	// Test AddAttachment and AddRelatedBodyPart with mock files
	importFilepath1 := "test_attach.txt"
	importFilepath2 := "test_related.txt"

	err1 := os.WriteFile(importFilepath1, []byte("attachment contents"), 0644)
	err2 := os.WriteFile(importFilepath2, []byte("related contents"), 0644)
	if err1 != nil || err2 != nil {
		t.Fatalf("failed to write mock files: %v, %v", err1, err2)
	}
	defer func() {
		os.Remove(importFilepath1)
		os.Remove(importFilepath2)
	}()

	// AddAttachment
	attachRet := obj.DispatchMethod("AddAttachment", []Value{NewString(importFilepath1)})
	if attachRet.Type != VTBool || attachRet.Num == 0 {
		t.Errorf("AddAttachment failed, got ret: %v", attachRet)
	}
	if len(obj.attachments) != 1 || obj.attachments[0] != importFilepath1 {
		t.Errorf("expected attachments to contain %s", importFilepath1)
	}

	// AddRelatedBodyPart
	bpVal := obj.DispatchMethod("AddRelatedBodyPart", []Value{NewString(importFilepath2), NewString("my-cid-123")})
	if bpVal.Type != VTNativeObject {
		t.Fatalf("expected VTNativeObject for related body part, got %v", bpVal.Type)
	}

	bpObj, ok := vm.g3mailItems[bpVal.Num]
	if !ok || bpObj.kind != 1 {
		t.Fatalf("expected kind 1 (body part) object, got %#v", bpObj)
	}

	// Test .Fields on body part
	fieldsVal := bpObj.DispatchPropertyGet("fields")
	if fieldsVal.Type != VTNativeObject {
		t.Fatalf("expected fields object to be VTNativeObject, got %v", fieldsVal.Type)
	}

	fieldsObj, ok := vm.g3mailItems[fieldsVal.Num]
	if !ok || fieldsObj.kind != 2 {
		t.Fatalf("expected kind 2 (fields collection) object, got %#v", fieldsObj)
	}

	// Test .Fields.Item("urn:schemas:mailheader:Content-ID") = "<my-cid-123>"
	// In JScript/VM execution this is called on the fields collection object via proxy or direct DispatchMethod
	setRet := fieldsObj.DispatchMethod("Item", []Value{NewString("urn:schemas:mailheader:Content-ID"), NewString("<my-new-cid>")})
	if setRet.Type != VTBool || setRet.Num == 0 {
		t.Errorf("Fields.Item set failed")
	}

	// Verify the CID was cleaned and updated in the body part
	if bpObj.cid != "my-new-cid" {
		t.Errorf("expected updated cid 'my-new-cid', got '%s'", bpObj.cid)
	}

	// Verify Fields.Item("urn...") getter
	getRet := fieldsObj.DispatchMethod("Item", []Value{NewString("urn:schemas:mailheader:Content-ID")})
	if getRet.String() != "<my-new-cid>" {
		t.Errorf("expected Fields.Item get to return '<my-new-cid>', got '%s'", getRet.String())
	}

	// Verify Fields.Update()
	updRet := fieldsObj.DispatchMethod("Update", nil)
	if updRet.Type != VTBool || updRet.Num == 0 {
		t.Errorf("Fields.Update failed")
	}
}
