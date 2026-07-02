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

// TestGlobalASAJScriptApplicationOnStart verifies that a JScript function
// Application_OnStart inside <script language="JScript" runat="server"> inside
// global.asa can set Application variables.  This exercises the JS env fallback
// path in executeHandler — the function is stored as a VTJSFunction in the JS
// environment, not as a VTUserSub in VBScript Globals.
func TestGlobalASAJScriptApplicationOnStart(t *testing.T) {
	tempDir := t.TempDir()
	asaPath := filepath.Join(tempDir, "global.asa")

	asaCode := `<script language="JScript" runat="server">
function Application_OnStart() {
    Application("IsStarted") = true;
    Application("Answer") = 42;
}
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

	// Verify Application("IsStarted") was set by the JScript handler.
	val, ok := app.Get("isstarted")
	if !ok {
		t.Fatal("expected Application(\"IsStarted\") to be set after JScript Application_OnStart")
	}
	if val.Type != asp.ApplicationValueBool || val.Num == 0 {
		t.Fatalf("expected Application(\"IsStarted\") to be true, got %#v", val)
	}

	// Verify Application("Answer") was set to 42.
	val, ok = app.Get("answer")
	if !ok {
		t.Fatal("expected Application(\"Answer\") to be set after JScript Application_OnStart")
	}
	if val.Type != asp.ApplicationValueInteger || val.Num != 42 {
		t.Fatalf("expected Application(\"Answer\") to be 42, got %#v", val)
	}
}

// TestGlobalASAJScriptSessionOnStart verifies that a JScript Session_OnStart
// function can set Session variables from global.asa.
func TestGlobalASAJScriptSessionOnStart(t *testing.T) {
	tempDir := t.TempDir()
	asaPath := filepath.Join(tempDir, "global.asa")

	asaCode := `<script language="JScript" runat="server">
function Session_OnStart() {
    Session("Language") = "JScript";
    Session("Score") = 99;
}
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

	session := asp.NewSession()
	host := NewMockHost()
	host.SetSession(session)
	host.SetApplication(app)
	host.SetOutput(new(bytes.Buffer))

	if err := globalASA.ExecuteSessionOnStart(host); err != nil {
		t.Fatalf("ExecuteSessionOnStart failed: %v", err)
	}

	// Verify Session("Language") was set by the JScript handler.
	val, ok := session.Get("language")
	if !ok {
		t.Fatal("expected Session(\"Language\") to be set after JScript Session_OnStart")
	}
	if val.Type != asp.ApplicationValueString || val.Str != "JScript" {
		t.Fatalf("expected Session(\"Language\") to be \"JScript\", got %#v", val)
	}

	// Verify Session("Score") was set to 99.
	val, ok = session.Get("score")
	if !ok {
		t.Fatal("expected Session(\"Score\") to be set after JScript Session_OnStart")
	}
	if val.Type != asp.ApplicationValueInteger || val.Num != 99 {
		t.Fatalf("expected Session(\"Score\") to be 99, got %#v", val)
	}
}

// TestGlobalASAJScriptApplicationOnStartNumericValue tests the exact
// reproduction case from the issue: Application("test") = 1 in JScript
// global.asa.
func TestGlobalASAJScriptApplicationOnStartNumericValue(t *testing.T) {
	tempDir := t.TempDir()
	asaPath := filepath.Join(tempDir, "global.asa")

	asaCode := `<script language="JScript" runat="server">
function Application_OnStart() {
    Application("test") = 1;
}
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

	// Verify Application("test") was set to 1.
	val, ok := app.Get("test")
	if !ok {
		t.Fatal("expected Application(\"test\") to be set after JScript Application_OnStart")
	}
	if val.Type != asp.ApplicationValueInteger || val.Num != 1 {
		t.Fatalf("expected Application(\"test\") to be 1, got %#v", val)
	}
}

// TestGlobalASAJScriptApplicationOnStartReadBack verifies that an Application
// variable set by a JScript global.asa handler can be read back from a JScript
// .asp page context.  This simulates the full lifecycle: global.asa sets the
// value, then a JScript page reads it.
func TestGlobalASAJScriptApplicationOnStartReadBack(t *testing.T) {
	tempDir := t.TempDir()
	asaPath := filepath.Join(tempDir, "global.asa")

	asaCode := `<script language="JScript" runat="server">
function Application_OnStart() {
    Application("test") = 1;
}
</script>`

	if err := os.WriteFile(asaPath, []byte(asaCode), 0644); err != nil {
		t.Fatalf("failed to write global.asa: %v", err)
	}

	app := asp.NewApplication()
	globalASA := &GlobalASA{}
	if err := globalASA.LoadAndCompile(tempDir, app); err != nil {
		t.Fatalf("failed to load global.asa: %v", err)
	}

	// 1) Execute Application_OnStart (JScript) to set Application("test") = 1.
	{
		host := NewMockHost()
		host.SetApplication(app)
		host.SetOutput(new(bytes.Buffer))
		if err := globalASA.ExecuteApplicationOnStart(host); err != nil {
			t.Fatalf("ExecuteApplicationOnStart failed: %v", err)
		}
	}

	// 2) Simulate a JScript .asp page reading Application("test").
	//    We compile a minimal JScript snippet that reads the value and writes it to
	//    the response as a number, then verify the page output.
	aspSource := `<%@ Language="JScript" %>
<%
var val = Application("test");
Response.Write(val);
%>`

	// Write the .asp file so the compiler can resolve its path.
	aspPath := filepath.Join(tempDir, "test_readback.asp")
	if err := os.WriteFile(aspPath, []byte(aspSource), 0644); err != nil {
		t.Fatalf("failed to write test .asp: %v", err)
	}

	compiler := NewASPCompiler(aspSource)
	compiler.SetSourceName(aspPath)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("failed to compile JScript .asp page: %v", err)
	}

	var buf bytes.Buffer
	vm := AcquireVMFromCompiler(compiler)
	defer vm.Release()

	host2 := NewMockHost()
	host2.SetApplication(app)
	host2.SetOutput(&buf)
	vm.SetHost(host2)

	if err := vm.Run(); err != nil {
		t.Fatalf("JScript .asp page execution failed: %v", err)
	}

	output := buf.String()
	if output != "1" {
		t.Fatalf("expected page output '1', got %q", output)
	}
}
