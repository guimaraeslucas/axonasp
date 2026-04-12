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

	obj.DispatchMethod("AddAddress", []Value{NewString("test@example.com")})
	if len(obj.to) != 1 || obj.to[0] != "test@example.com" {
		t.Error("AddAddress failed")
	}
}
