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
 */
package axonvm

import (
	"bytes"
	"strings"
	"testing"
)

func TestJScriptPromise(t *testing.T) {
	script := `
		var result = "PENDING";
		var p = new Promise(function(resolve, reject) {
			resolve("SUCCESS");
		});
		p.then(function(val) {
			result = val;
			Response.Write(result);
		});
	`

	compiler := NewASPCompiler("<%@ Language=\"JavaScript\" %>\n<% " + script + " %>")
	if err := compiler.Compile(); err != nil {
		t.Fatalf("Compile failed: %v", err)
	}
	vm := AcquireVMFromCompiler(compiler)
	defer vm.Release()
	host := NewMockHost()
	var out bytes.Buffer
	host.SetOutput(&out)
	host.Response().SetBuffer(false)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("Run failed: %v", err)
	}

	output := out.String()
	if !strings.Contains(output, "SUCCESS") {
		t.Errorf("Expected output to contain 'SUCCESS', got '%s'", output)
	}
}

func TestJScriptPromiseAll(t *testing.T) {
	script := `
		var p1 = Promise.resolve(1);
		var p2 = Promise.resolve(2);
		Promise.all([p1, p2]).then(function(values) {
			Response.Write(values.join(","));
		});
	`

	compiler := NewASPCompiler("<%@ Language=\"JavaScript\" %>\n<% " + script + " %>")
	if err := compiler.Compile(); err != nil {
		t.Fatalf("Compile failed: %v", err)
	}
	vm := AcquireVMFromCompiler(compiler)
	defer vm.Release()
	host := NewMockHost()
	var out bytes.Buffer
	host.SetOutput(&out)
	host.Response().SetBuffer(false)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("Run failed: %v", err)
	}

	output := out.String()
	if !strings.Contains(output, "1,2") {
		t.Errorf("Expected output to contain '1,2', got '%s'", output)
	}
}
