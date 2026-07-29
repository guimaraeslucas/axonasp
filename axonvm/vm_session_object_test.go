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
	"testing"
)

// TestASPSessionAndApplicationObjectReferences verifies storing, retrieving,
// and invoking methods on native objects (e.g., Scripting.Dictionary) via Session and Application.
func TestASPSessionAndApplicationObjectReferences(t *testing.T) {
	source := `<%
	Dim dict
	Set dict = Server.CreateObject("Scripting.Dictionary")
	dict.Add "title", "Hello World"
	dict.Add "count", 42

	Response.Write "Step1: " & TypeName(dict) & "|" & dict.Count & "|" & dict("title") & ";"

	Set Session("myDict") = dict

	Dim retrievedSess
	Set retrievedSess = Session("myDict")

	Response.Write "Step2: " & TypeName(retrievedSess) & "|" & CStr(IsObject(retrievedSess)) & "|" & retrievedSess.Count & ";"

	Dim c1
	c1 = retrievedSess.Count
	Response.Write "Step3: " & c1 & ";"

	Set Application("myAppDict") = dict
	Dim retrievedApp
	Set retrievedApp = Application("myAppDict")

	Response.Write "Step4: " & TypeName(retrievedApp) & "|" & CStr(IsObject(retrievedApp)) & "|" & retrievedApp.Count & ";"
	%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	expected := "Step1: Dictionary|2|Hello World;Step2: Dictionary|True|2;Step3: 2;Step4: Dictionary|True|2;"
	if output.String() != expected {
		t.Fatalf("expected output %q, got %q", expected, output.String())
	}
}
