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

func TestJScriptEvalSyntaxErrorCatchable(t *testing.T) {
	source := `<%@ Language="JScript" CodePage="65001" %>
<%
var errObj;
try {
    eval("x = @@@");   // malformed source
} catch(e) {
    errObj = e;        // IIS catches this.
}
Response.Write("caught: " + (errObj ? errObj.message : "(none)"));
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("top-level compile failed unexpectedly: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	host.Response().SetBuffer(false)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed unexpectedly: %v", err)
	}

	got := output.String()
	t.Logf("Output: %s", got)
	if !strings.Contains(got, "caught: Syntax error") {
		t.Fatalf("expected 'caught: Syntax error', got: %s", got)
	}
}

func TestJScriptNewFunctionSyntaxErrorCatchable(t *testing.T) {
	source := `<%@ Language="JScript" CodePage="65001" %>
<%
var errObj;
try {
    new Function("x = @@@");
} catch(e) {
    errObj = e;
}
Response.Write("caught: " + (errObj ? errObj.message : "(none)"));
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("top-level compile failed unexpectedly: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	host.Response().SetBuffer(false)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed unexpectedly: %v", err)
	}

	got := output.String()
	t.Logf("Output: %s", got)
	if !strings.Contains(got, "caught: Syntax error") {
		t.Fatalf("expected 'caught: Syntax error', got: %s", got)
	}
}

func TestJScriptTopLevelSyntaxErrorFatal(t *testing.T) {
	source := `<%@ Language="JScript" CodePage="65001" %>
<%
var x = @@@;
%>`

	compiler := NewASPCompiler(source)
	err := compiler.Compile()
	if err == nil {
		t.Fatalf("expected top-level syntax error during compilation, but got no error")
	}
	t.Logf("Top-level compilation correctly failed: %v", err)
}
