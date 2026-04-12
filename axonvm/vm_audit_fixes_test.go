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
	"path/filepath"
	"strconv"
	"strings"
	"testing"

	"g3pix.com.br/axonasp/vbscript"
)

func runASPAndCollectOutput(t *testing.T, source string) string {
	t.Helper()

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	host.Server().SetRootDir(t.TempDir())
	host.Server().SetRequestPath("/tests/audit.asp")

	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()
	return output.String()
}

func TestASPSetSupportsChainedCallMemberAssignment(t *testing.T) {
	source := `<%
Class Node
	Public Child
End Class

Function Pick(v)
	Set Pick = v
End Function

Dim a
Dim b
Set a = New Node
Set b = New Node
Set a.Child = b

On Error Resume Next
Set Pick(a).Child.Child = New Node
Response.Write Err.Number & "|" & TypeName(a.Child.Child)
%>`

	actual := runASPAndCollectOutput(t, source)
	if actual != "0|Node" {
		t.Fatalf("unexpected output: got %q want %q", actual, "0|Node")
	}
}

func TestASPMemberGetOnNothingRaisesObjectRequired424(t *testing.T) {
	source := `<%
On Error Resume Next
Dim o
Set o = Nothing
Dim v
v = o.Name
Response.Write Err.Number
%>`

	actual := runASPAndCollectOutput(t, source)
	expected := strconv.Itoa(vbscript.HRESULTFromVBScriptCode(vbscript.CouldNotFindTargetObject))
	if actual != expected {
		t.Fatalf("unexpected Err.Number: got %q want %q", actual, expected)
	}
}

func TestASPMemberCallOnScalarRaisesObjectRequired424(t *testing.T) {
	source := `<%
On Error Resume Next
Dim n
n = 1
Call n.DoWork()
Response.Write Err.Number
%>`

	actual := runASPAndCollectOutput(t, source)
	expected := strconv.Itoa(vbscript.HRESULTFromVBScriptCode(vbscript.CouldNotFindTargetObject))
	if actual != expected {
		t.Fatalf("unexpected Err.Number: got %q want %q", actual, expected)
	}
}

func TestVMADODBConnectionErrorsCollectionPopulatesOnFailure(t *testing.T) {
	vm := NewVM(nil, nil, 5)
	host := NewMockHost()
	rootDir := t.TempDir()
	host.Server().SetRootDir(rootDir)
	host.Server().SetRequestPath("/tests/adodb-errors.asp")
	vm.SetHost(host)

	conn := vm.dispatchNativeCall(nativeObjectServer, "CreateObject", []Value{NewString("ADODB.Connection")})
	if conn.Type != VTNativeObject {
		t.Fatalf("expected ADODB.Connection native object, got %#v", conn)
	}

	dbPath := filepath.Join(rootDir, "audit-errors.db")
	vm.dispatchMemberSet(conn.Num, "ConnectionString", NewString("sqlite:"+dbPath))
	vm.dispatchNativeCall(conn.Num, "Open", nil)
	defer vm.dispatchNativeCall(conn.Num, "Close", nil)

	vm.onResumeNext = true
	vm.dispatchNativeCall(conn.Num, "Execute", []Value{NewString("SELECT * FROM __audit_missing_table__")})

	errorsCollection := vm.dispatchMemberGet(conn, "Errors")
	if errorsCollection.Type != VTNativeObject {
		t.Fatalf("expected Errors collection object, got %#v", errorsCollection)
	}

	count := vm.dispatchMemberGet(errorsCollection, "Count")
	if count.Type != VTInteger || count.Num < 1 {
		t.Fatalf("expected Errors.Count >= 1, got %#v", count)
	}

	item := vm.dispatchNativeCall(errorsCollection.Num, "Item", []Value{NewInteger(0)})
	if item.Type != VTNativeObject {
		t.Fatalf("expected first Errors.Item to be an object, got %#v", item)
	}

	description := vm.dispatchMemberGet(item, "Description")
	if description.Type != VTString || strings.TrimSpace(description.Str) == "" {
		t.Fatalf("expected non-empty error description, got %#v", description)
	}

	expectedErr := vbscript.HRESULTFromVBScriptCode(vbscript.AutomationError)
	if number := vm.errPropertyValue("Number"); number.Type != VTInteger || int(number.Num) != expectedErr {
		t.Fatalf("unexpected Err.Number after Execute failure: got %#v want %d", number, expectedErr)
	}
}

func TestVMADODBConnectionOpenEmptyStringRaisesAutomationError(t *testing.T) {
	vm := NewVM(nil, nil, 5)
	host := NewMockHost()
	host.Server().SetRootDir(t.TempDir())
	host.Server().SetRequestPath("/tests/adodb-open-empty.asp")
	vm.SetHost(host)

	conn := vm.dispatchNativeCall(nativeObjectServer, "CreateObject", []Value{NewString("ADODB.Connection")})
	if conn.Type != VTNativeObject {
		t.Fatalf("expected ADODB.Connection native object, got %#v", conn)
	}

	vm.onResumeNext = true
	vm.dispatchNativeCall(conn.Num, "Open", nil)

	expectedErr := vbscript.HRESULTFromVBScriptCode(vbscript.AutomationError)
	if number := vm.errPropertyValue("Number"); number.Type != VTInteger || int(number.Num) != expectedErr {
		t.Fatalf("unexpected Err.Number after Open failure: got %#v want %d", number, expectedErr)
	}
}

func TestVMFSOGetFileAndDeleteFolderRaiseClassicErrors(t *testing.T) {
	vm := NewVM(nil, nil, 5)
	host := NewMockHost()
	host.Server().SetRootDir(t.TempDir())
	host.Server().SetRequestPath("/tests/fso-errors.asp")
	vm.SetHost(host)

	fso := vm.dispatchNativeCall(nativeObjectServer, "CreateObject", []Value{NewString("Scripting.FileSystemObject")})
	if fso.Type != VTNativeObject {
		t.Fatalf("expected Scripting.FileSystemObject native object, got %#v", fso)
	}

	vm.onResumeNext = true
	vm.dispatchNativeCall(fso.Num, "GetFile", []Value{NewString("/missing-file.txt")})
	expectedFileNotFound := vbscript.HRESULTFromVBScriptCode(vbscript.FileNotFound)
	if number := vm.errPropertyValue("Number"); number.Type != VTInteger || int(number.Num) != expectedFileNotFound {
		t.Fatalf("unexpected Err.Number for GetFile missing path: got %#v want %d", number, expectedFileNotFound)
	}

	vm.errClear()
	vm.dispatchNativeCall(fso.Num, "DeleteFolder", []Value{NewString("/missing-folder")})
	expectedPathNotFound := vbscript.HRESULTFromVBScriptCode(vbscript.PathNotFound)
	if number := vm.errPropertyValue("Number"); number.Type != VTInteger || int(number.Num) != expectedPathNotFound {
		t.Fatalf("unexpected Err.Number for DeleteFolder missing path: got %#v want %d", number, expectedPathNotFound)
	}
}
