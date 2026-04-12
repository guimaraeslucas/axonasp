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
	"os"
	"path/filepath"
	"testing"
)

// TestVMForEachIteratesFSOFilesCollection verifies For Each can iterate Scripting.FileSystemObject Files collections.
func TestVMForEachIteratesFSOFilesCollection(t *testing.T) {
	tempDir := t.TempDir()
	for _, name := range []string{"b.txt", "a.txt", "c.txt"} {
		if err := os.WriteFile(filepath.Join(tempDir, name), []byte("ok"), 0644); err != nil {
			t.Fatalf("write sample file %s: %v", name, err)
		}
	}

	source := `<%
Dim fso, folderObj, fileObj, fileCount
Set fso = Server.CreateObject("Scripting.FileSystemObject")
Set folderObj = fso.GetFolder(Server.MapPath("/"))
fileCount = 0
Response.Write folderObj.Files.Count
Response.Write ":"
For Each fileObj In folderObj.Files
    fileCount = fileCount + 1
    Response.Write fileObj.Name & ";"
Next
Response.Write ":" & fileCount
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	host.Server().SetRootDir(tempDir)
	host.Server().SetRequestPath("/index.asp")
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	if output.String() != "3:a.txt;b.txt;c.txt;:3" {
		t.Fatalf("unexpected file iteration output: got %q want %q", output.String(), "3:a.txt;b.txt;c.txt;:3")
	}
}

// TestVMForEachIteratesFSOSubFoldersCollection verifies SubFolders enumeration is alphabetical.
func TestVMForEachIteratesFSOSubFoldersCollection(t *testing.T) {
	tempDir := t.TempDir()
	for _, name := range []string{"beta", "alpha"} {
		if err := os.Mkdir(filepath.Join(tempDir, name), 0755); err != nil {
			t.Fatalf("create subfolder %s: %v", name, err)
		}
	}

	source := `<%
Dim fso, folderObj, subFolderObj, folderCount
Set fso = Server.CreateObject("Scripting.FileSystemObject")
Set folderObj = fso.GetFolder(Server.MapPath("/"))
folderCount = 0
Response.Write folderObj.SubFolders.Count
Response.Write ":"
For Each subFolderObj In folderObj.SubFolders
    folderCount = folderCount + 1
    Response.Write subFolderObj.Name & ";"
Next
Response.Write ":" & folderCount
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	host.Server().SetRootDir(tempDir)
	host.Server().SetRequestPath("/index.asp")
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	if output.String() != "2:alpha;beta;:2" {
		t.Fatalf("unexpected subfolder iteration output: got %q want %q", output.String(), "2:alpha;beta;:2")
	}
}

// TestVMFunctionOptionalParameter verifies Optional parameter declarations compile and default to Empty values.
func TestVMFunctionOptionalParameter(t *testing.T) {
	source := `<%
Function asperror(Optional path)
    If Len(path) = 0 Then
        Response.Write "EMPTY"
    Else
        Response.Write path
    End If
End Function

result = asperror()
Response.Write "|"
result = asperror("/tests")
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

	if output.String() != "EMPTY|/tests" {
		t.Fatalf("unexpected Optional behavior output: got %q want %q", output.String(), "EMPTY|/tests")
	}
}

// TestVMExpandedVBScriptConstants verifies key VBScript constants map to classic expected values.
func TestVMExpandedVBScriptConstants(t *testing.T) {
	source := `<%= vbJanuary %>|<%= vbObjectError %>|<%= vbOKOnly %>|<%= vbCritical %>|<%= vbUseSystem %>|<%= vbASCII %>`

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

	expected := "1|-2147221504|0|16|0|0"
	if output.String() != expected {
		t.Fatalf("unexpected constants output: got %q want %q", output.String(), expected)
	}
}
