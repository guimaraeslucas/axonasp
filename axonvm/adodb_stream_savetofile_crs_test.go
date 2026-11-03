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

func TestADODBStreamSaveToFileRewindsPosition(t *testing.T) {
	src := "<%\nSet s = Server.CreateObject(\"ADODB.Stream\") : s.Charset = \"UTF-8\" : s.Type = 2 : s.Open\ns.WriteText \"row1\" : s.WriteText \"row2\"\ns.SaveToFile Server.MapPath(\"/_axon_stream_test.txt\"), 2\nResponse.Write s.Position & \":\" & s.ReadText()\n%>"
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
	if got := output.String(); got != "0:row1row2" {
		t.Fatalf("got %q, want %q", got, "0:row1row2")
	}
}
