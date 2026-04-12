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
	"os"
	"path/filepath"
	"testing"

	"g3pix.com.br/axonasp/axonvm/asp"
)

func TestGlobalASAApplicationOnStart(t *testing.T) {
	tempDir := t.TempDir()
	asaPath := filepath.Join(tempDir, "global.asa")

	asaCode := `<script runat="server" language="VBScript">
Sub Application_OnStart
    Application("IsStarted") = True
End Sub
</script>`

	if err := os.WriteFile(asaPath, []byte(asaCode), 0644); err != nil {
		t.Fatalf("failed to write global.asa: %v", err)
	}

	app := asp.NewApplication()
	globalASA := &GlobalASA{}
	if err := globalASA.LoadAndCompile(tempDir, app); err != nil {
		t.Fatalf("failed to load global.asa: %v", err)
	}

	if !globalASA.IsLoaded() {
		t.Fatal("expected global.asa to be marked as loaded")
	}

	host := NewMockHost()
	host.SetApplication(app)
	host.SetOutput(new(bytes.Buffer))

	if err := globalASA.ExecuteApplicationOnStart(host); err != nil {
		t.Fatalf("ExecuteApplicationOnStart failed: %v", err)
	}

	val, ok := app.Get("isstarted")
	if !ok {
		t.Fatalf("expected Application(\"IsStarted\") to be set")
	}

	if val.Type != asp.ApplicationValueBool || val.Num == 0 {
		t.Fatalf("expected Application(\"IsStarted\") to be True, got %#v", val)
	}
}

func TestGlobalASAObjectDeclarations(t *testing.T) {
	tempDir := t.TempDir()
	asaPath := filepath.Join(tempDir, "global.asa")

	asaCode := `<object runat="server" scope="Application" id="AppObj" progid="Scripting.Dictionary"></object>
<object runat="server" scope="Session" id="SessObj" progid="Scripting.FileSystemObject"></object>`

	if err := os.WriteFile(asaPath, []byte(asaCode), 0644); err != nil {
		t.Fatalf("failed to write global.asa: %v", err)
	}

	app := asp.NewApplication()
	globalASA := &GlobalASA{}
	if err := globalASA.LoadAndCompile(tempDir, app); err != nil {
		t.Fatalf("failed to load global.asa: %v", err)
	}

	// Verify Application StaticObject
	if !app.ContainsStaticObject("appobj") {
		t.Fatal("expected Application to contain AppObj static object")
	}

	// Verify Session StaticObject is populated upon new session creation
	session := asp.NewSession()
	globalASA.PopulateSessionStaticObjects(session)

	if !session.ContainsStaticObject("sessobj") {
		t.Fatal("expected Session to contain SessObj static object")
	}
}
