package axonvm

/*
 * AxonASP Server
 * Copyright (C) 2026 G3pix Ltda. All rights reserved.
 *
 * Developed by Yuri Eisman (@yeisman)
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

import (
	"bytes"
	"testing"
)

func TestClassFieldDictionaryDefaultPropertySet(t *testing.T) {
	src := "<%\nClass C\n Public data\n Private Sub Class_Initialize : Set data = CreateObject(\"Scripting.Dictionary\") : End Sub\nEnd Class\nSet o = New C\no.data(\"a\") = 1\no.data(\"a\") = o.data(\"a\") + 1\no.data(\"b\") = \"x\"\nResponse.Write o.data(\"a\") & o.data(\"b\") & o.data.Count\n%>"
	compiler := NewASPCompiler(src)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}
	vm := NewVMFromCompiler(compiler)
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)
	if err := vm.Run(); err != nil {
		t.Fatalf("run failed: %v", err)
	}
	if got := output.String(); got != "2x2" {
		t.Fatalf("got %q, want %q", got, "2x2")
	}
}
