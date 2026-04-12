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
	"strings"
	"testing"
)

func TestMSWC(t *testing.T) {
	vm := NewVM(nil, nil, 0)
	host := NewMockHost()
	host.Server().SetRootDir(".")
	vm.host = host

	// Test BrowserType
	bt := vm.newG3BrowserTypeObject()
	objBT := vm.mswcBrowserTypeItems[bt.Num]
	if objBT == nil {
		t.Fatal("expected G3BrowserType")
	}
	browser := objBT.DispatchPropertyGet("browser")
	if browser.String() != "Unknown" {
		t.Errorf("expected Unknown, got %s", browser.String())
	}

	// Test Counters
	ct := vm.newG3CountersObject()
	objCT := vm.mswcCountersItems[ct.Num]

	objCT.DispatchMethod("Set", []Value{NewString("hits"), NewInteger(100)})
	val := objCT.DispatchMethod("Get", []Value{NewString("hits")})
	if val.Num != 100 {
		t.Errorf("expected 100, got %d", val.Num)
	}

	// Test Tools
	tls := vm.newG3ToolsObject()
	objTls := vm.mswcToolsItems[tls.Num]

	// create dummy file to test owner/exists
	f, _ := os.Create("test.txt")
	f.Close()
	defer os.Remove("test.txt")

	exists := objTls.DispatchMethod("FileExists", []Value{NewString("test.txt")})
	if exists.Type != VTBool {
		t.Error("expected bool from FileExists")
	}

	// Test PermissionChecker
	pcValue := vm.newG3PermissionCheckerObject()
	objPC := vm.mswcPermissionCheckerItems[pcValue.Num]
	if objPC == nil {
		t.Fatal("expected G3PermissionChecker")
	}

	// Test valid file
	hasAccess := objPC.DispatchMethod("HasAccess", []Value{NewString("test.txt")})
	if hasAccess.Type != VTBool || hasAccess.Num == 0 {
		t.Error("expected HasAccess to return true for existing file")
	}

	// Test non-existent file
	noAccess := objPC.DispatchMethod("HasAccess", []Value{NewString("nonexistent.txt")})
	if noAccess.Type != VTBool || noAccess.Num != 0 {
		t.Error("expected HasAccess to return false for non-existent file")
	}
}

func TestPageCounter(t *testing.T) {
	vm := NewVM(nil, nil, 0)

	// The default in config/axonasp.toml is disabled (false).
	// Let's verify that instantiation raises an error.
	defer func() {
		if r := recover(); r != nil {
			vme, ok := r.(*VMError)
			if !ok {
				t.Fatalf("expected VMError, got %T", r)
			}
			msg := vme.Error()
			if !strings.Contains(msg, "MSWC.PageCounter is disabled") {
				t.Errorf("expected disabled error message, got: %s", msg)
			}
		} else {
			t.Error("expected panic when PageCounter is disabled")
		}
	}()

	_ = vm.newG3PageCounterObject()
}

func TestPageCounterEnabled(t *testing.T) {
	// This test depends on environment variables being set before execution
	if os.Getenv("MSWC_PAGECOUNTER_ENABLED") != "true" {
		t.Skip("Skipping TestPageCounterEnabled; environment variable MSWC_PAGECOUNTER_ENABLED=true not set")
	}

	vm := NewVM(nil, nil, 0)
	pc := vm.newG3PageCounterObject()
	obj := vm.mswcPageCounterItems[pc.Num]

	// Simulate hit
	host := NewMockHost()
	host.Request().ServerVars.Add("SCRIPT_NAME", "/test.asp")
	vm.host = host

	val := obj.DispatchMethod("PageHit", nil)
	if val.Num != 1 {
		t.Errorf("expected 1 hit, got %d", val.Num)
	}

	// Verify hits method
	hits := obj.DispatchMethod("Hits", []Value{NewString("/test.asp")})
	if hits.Num != 1 {
		t.Errorf("expected 1 hit from Hits method, got %d", hits.Num)
	}
}
