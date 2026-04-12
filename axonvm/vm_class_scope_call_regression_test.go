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
	"bytes"
	"testing"
)

// TestVMClassMethodCallShadowsGlobal verifies that an unqualified class method
// call in statement position binds to the current class member even when a
// global variable with the same name exists.
func TestVMClassMethodCallShadowsGlobal(t *testing.T) {
	source := `<%
Dim Upload
Class U
	Public Value
	Private Sub Class_Initialize()
		Value = 0
	End Sub
	Public Sub Save()
		If Not False Then Upload
	End Sub
	Public Sub Upload()
		Value = 1
	End Sub
End Class
Set Upload = New U
Upload.Save
Response.Write Upload.Value
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

	if output.String() != "1" {
		t.Fatalf("expected class Upload method to be invoked, got output %q", output.String())
	}
}
