//go:build !wasm

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
	"errors"
	"os"
	"path/filepath"
	"runtime"
	"strings"
	"sync"
	"testing"

	"g3pix.com.br/axonasp/axonvm/asp"
)

// TestVMServerEncodingDispatch verifies native Server encoding method dispatch.
func TestVMServerEncodingDispatch(t *testing.T) {
	vm := NewVM(nil, nil, 5)
	host := NewMockHost()
	vm.SetHost(host)

	html := vm.dispatchNativeCall(2, "HTMLEncode", []Value{NewString("<x>")})
	if html.Type != VTString || html.Str != "&lt;x&gt;" {
		t.Fatalf("unexpected HTMLEncode result: %#v", html)
	}

	url := vm.dispatchNativeCall(2, "URLEncode", []Value{NewString("a b")})
	if url.Type != VTString || url.Str != "a+b" {
		t.Fatalf("unexpected URLEncode result: %#v", url)
	}
}

// TestVMServerFSOTimestampProperties verifies File and Folder date properties use platform-accurate timestamps.
func TestVMServerFSOTimestampProperties(t *testing.T) {
	vm := NewVM(nil, nil, 5)
	host := NewMockHost()
	rootDir := t.TempDir()
	host.Server().SetRootDir(rootDir)
	host.Server().SetRequestPath("/tests/test_server.asp")
	vm.SetHost(host)

	if err := os.MkdirAll(filepath.Join(rootDir, "timestamps"), 0755); err != nil {
		t.Fatalf("mkdir timestamps: %v", err)
	}
	filePath := filepath.Join(rootDir, "timestamps", "sample.txt")
	if err := os.WriteFile(filePath, []byte("timestamps"), 0644); err != nil {
		t.Fatalf("write sample file: %v", err)
	}

	fso := vm.dispatchNativeCall(nativeObjectServer, "CreateObject", []Value{NewString("Scripting.FileSystemObject")})
	if fso.Type != VTNativeObject {
		t.Fatalf("expected VTNativeObject for FSO, got %#v", fso)
	}

	fileObj := vm.dispatchNativeCall(fso.Num, "GetFile", []Value{NewString("/timestamps/sample.txt")})
	folderObj := vm.dispatchNativeCall(fso.Num, "GetFolder", []Value{NewString("/timestamps")})
	if fileObj.Type != VTNativeObject || folderObj.Type != VTNativeObject {
		t.Fatalf("expected file and folder objects, got file=%#v folder=%#v", fileObj, folderObj)
	}

	fileCreated := vm.dispatchMemberGet(fileObj, "DateCreated")
	fileAccessed := vm.dispatchMemberGet(fileObj, "DateLastAccessed")
	fileModified := vm.dispatchMemberGet(fileObj, "DateLastModified")
	folderCreated := vm.dispatchMemberGet(folderObj, "DateCreated")
	folderAccessed := vm.dispatchMemberGet(folderObj, "DateLastAccessed")
	folderModified := vm.dispatchMemberGet(folderObj, "DateLastModified")

	dateValues := []Value{fileCreated, fileAccessed, fileModified, folderCreated, folderAccessed, folderModified}
	for i := range dateValues {
		if dateValues[i].Type != VTDate {
			t.Fatalf("expected VTDate value at index %d, got %#v", i, dateValues[i])
		}
	}

	if runtime.GOOS != "windows" {
		if fileCreated.Num != fileModified.Num || fileAccessed.Num != fileModified.Num {
			t.Fatalf("expected non-Windows file timestamps to fallback to ModTime, got created=%#v accessed=%#v modified=%#v", fileCreated, fileAccessed, fileModified)
		}
		if folderCreated.Num != folderModified.Num || folderAccessed.Num != folderModified.Num {
			t.Fatalf("expected non-Windows folder timestamps to fallback to ModTime, got created=%#v accessed=%#v modified=%#v", folderCreated, folderAccessed, folderModified)
		}
	}
}

// TestVMServerMapPathAndTimeout verifies MapPath and ScriptTimeout dispatch.
func TestVMServerMapPathAndTimeout(t *testing.T) {
	vm := NewVM(nil, nil, 5)
	host := NewMockHost()
	host.Server().SetRootDir("./www")
	host.Server().SetRequestPath("/tests/page.asp")
	vm.SetHost(host)

	mapped := vm.dispatchNativeCall(2, "MapPath", []Value{NewString("local.asp")})
	if mapped.Type != VTString || !strings.Contains(strings.ReplaceAll(mapped.Str, "\\", "/"), "/www/tests/local.asp") {
		t.Fatalf("unexpected mapped path: %#v", mapped)
	}

	connected := vm.dispatchNativeCall(2, "IsClientConnected", nil)
	if connected.Type != VTBool || connected.Num != 1 {
		t.Fatalf("unexpected IsClientConnected value: %#v", connected)
	}

	vm.dispatchNativeCall(2, "ScriptTimeout", []Value{NewInteger(123)})
	timeout := vm.dispatchNativeCall(2, "ScriptTimeout", nil)
	if timeout.Type != VTInteger || timeout.Num != 123 {
		t.Fatalf("unexpected timeout value: %#v", timeout)
	}

	vm.dispatchMemberSet(nativeObjectServer, "ScriptTimeout", NewInteger(77))
	timeout = vm.dispatchMemberGet(Value{Type: VTNativeObject, Num: nativeObjectServer}, "ScriptTimeout")
	if timeout.Type != VTInteger || timeout.Num != 77 {
		t.Fatalf("unexpected property timeout value: %#v", timeout)
	}
}

// TestVMServerGetLastError verifies VM access to last error properties.
func TestVMServerGetLastError(t *testing.T) {
	vm := NewVM(nil, nil, 5)
	host := NewMockHost()
	vm.SetHost(host)

	// Use an unsupported ProgID to guarantee Server.CreateObject sets LastError.
	vm.onResumeNext = true
	vm.dispatchNativeCall(2, "CreateObject", []Value{NewString("AxonASP.Unknown.Component")})
	vm.onResumeNext = false

	number := vm.dispatchNativeCall(2, "GetLastError", []Value{NewString("Number")})
	if number.Type != VTInteger || number.Num != int64(asp.InvalidProgIDHRESULT) {
		t.Fatalf("unexpected error number: %#v", number)
	}

	description := vm.dispatchNativeCall(2, "GetLastError", []Value{NewString("Description")})
	if description.Type != VTString || description.Str == "" {
		t.Fatalf("unexpected error description: %#v", description)
	}

	obj := vm.dispatchNativeCall(2, "GetLastError", nil)
	if obj.Type != VTNativeObject {
		t.Fatalf("expected ASPError native object, got %#v", obj)
	}

	source := vm.dispatchMemberGet(obj, "Source")
	if source.Type != VTString || source.Str != "Server.CreateObject" {
		t.Fatalf("unexpected ASPError source: %#v", source)
	}

	aspCode := vm.dispatchMemberGet(obj, "ASPCode")
	if aspCode.Type != VTInteger || aspCode.Num != 429 {
		t.Fatalf("unexpected ASPError ASPCode: %#v", aspCode)
	}
}

// TestVMServerCreateObjectRegExpAlias verifies Server.CreateObject supports RegExp aliases.
func TestVMServerCreateObjectRegExpAlias(t *testing.T) {
	vm := NewVM(nil, nil, 5)
	host := NewMockHost()
	vm.SetHost(host)

	objA := vm.dispatchNativeCall(nativeObjectServer, "CreateObject", []Value{NewString("RegExp")})
	if objA.Type != VTNativeObject {
		t.Fatalf("expected RegExp alias to return native object, got %#v", objA)
	}

	objB := vm.dispatchNativeCall(nativeObjectServer, "CreateObject", []Value{NewString("VBScript.RegExp")})
	if objB.Type != VTNativeObject {
		t.Fatalf("expected VBScript.RegExp alias to return native object, got %#v", objB)
	}

	if objA.Num == objB.Num {
		t.Fatalf("expected distinct object IDs for each CreateObject call, got same ID %d", objA.Num)
	}
}

// TestVMErrASPCodeStringDispatch verifies intrinsic Err preserves non-numeric ASPCode values via dispatch.
func TestVMErrASPCodeStringDispatch(t *testing.T) {
	vm := NewVM(nil, nil, 9)
	host := NewMockHost()
	vm.SetHost(host)

	vm.dispatchMemberSet(nativeObjectErr, "ASPCode", NewString("ERR_TEST_001"))

	got := vm.dispatchMemberGet(Value{Type: VTNativeObject, Num: nativeObjectErr}, "ASPCode")
	if got.Type != VTString || got.Str != "ERR_TEST_001" {
		t.Fatalf("unexpected Err.ASPCode via dispatch: %#v", got)
	}
}

// TestVMServerCreateObjectG3MD verifies CreateObject for G3MD native library.
func TestVMServerCreateObjectG3MD(t *testing.T) {
	vm := NewVM(nil, nil, 5)
	host := NewMockHost()
	vm.SetHost(host)

	obj := vm.dispatchNativeCall(nativeObjectServer, "CreateObject", []Value{NewString("G3Md")})
	if obj.Type != VTNativeObject {
		t.Fatalf("expected VTNativeObject, got %#v", obj)
	}

	hardWrapsDefault := vm.dispatchMemberGet(obj, "HardWraps")
	if hardWrapsDefault.Type != VTBool || hardWrapsDefault.Num != 0 {
		t.Fatalf("unexpected HardWraps default: %#v", hardWrapsDefault)
	}

	unsafeDefault := vm.dispatchMemberGet(obj, "Unsafe")
	if unsafeDefault.Type != VTBool || unsafeDefault.Num != 0 {
		t.Fatalf("unexpected Unsafe default: %#v", unsafeDefault)
	}
}

// TestVMServerCreateObjectG3MDLowercase verifies lowercase ProgID resolution for G3MD.
func TestVMServerCreateObjectG3MDLowercase(t *testing.T) {
	vm := NewVM(nil, nil, 5)
	host := NewMockHost()
	vm.SetHost(host)

	obj := vm.dispatchNativeCall(nativeObjectServer, "CreateObject", []Value{NewString("g3md")})
	if obj.Type != VTNativeObject {
		t.Fatalf("expected VTNativeObject, got %#v", obj)
	}

	result := vm.dispatchNativeCall(obj.Num, "Process", []Value{NewString("# title")})
	if result.Type != VTString || !strings.Contains(strings.ToLower(result.Str), "title") {
		t.Fatalf("unexpected Process result: %#v", result)
	}
}

// TestVMServerG3MDProcess verifies G3MD property assignment and markdown conversion.
func TestVMServerG3MDProcess(t *testing.T) {
	vm := NewVM(nil, nil, 5)
	host := NewMockHost()
	vm.SetHost(host)

	obj := vm.dispatchNativeCall(nativeObjectServer, "CreateObject", []Value{NewString("G3MD")})

	vm.dispatchNativeCall(obj.Num, "HardWraps", []Value{NewBool(true)})
	vm.dispatchNativeCall(obj.Num, "Unsafe", []Value{NewInteger(1)})

	hardWraps := vm.dispatchMemberGet(obj, "HardWraps")
	if hardWraps.Type != VTBool || hardWraps.Num != 1 {
		t.Fatalf("unexpected HardWraps value: %#v", hardWraps)
	}

	unsafeValue := vm.dispatchMemberGet(obj, "Unsafe")
	if unsafeValue.Type != VTBool || unsafeValue.Num != 1 {
		t.Fatalf("unexpected Unsafe value: %#v", unsafeValue)
	}

	result := vm.dispatchNativeCall(obj.Num, "Process", []Value{NewString("line1\nline2")})
	if result.Type != VTString {
		t.Fatalf("expected VTString result, got %#v", result)
	}

	if !strings.Contains(result.Str, "line1") || !strings.Contains(result.Str, "line2") {
		t.Fatalf("unexpected Process output: %q", result.Str)
	}
}

// TestVMServerCreateObjectG3StringBuilder verifies CreateObject for G3STRINGBUILDER native library.
func TestVMServerCreateObjectG3StringBuilder(t *testing.T) {
	vm := NewVM(nil, nil, 5)
	host := NewMockHost()
	vm.SetHost(host)

	obj := vm.dispatchNativeCall(nativeObjectServer, "CreateObject", []Value{NewString("G3STRINGBUILDER")})
	if obj.Type != VTNativeObject {
		t.Fatalf("expected VTNativeObject, got %#v", obj)
	}

	vm.dispatchNativeCall(obj.Num, "Append", []Value{NewString("A")})
	vm.dispatchNativeCall(obj.Num, "append", []Value{NewInteger(2)})
	vm.dispatchNativeCall(obj.Num, "Append", []Value{NewString("C")})

	result := vm.dispatchNativeCall(obj.Num, "ToString", nil)
	if result.Type != VTString || result.Str != "A2C" {
		t.Fatalf("unexpected ToString result: %#v", result)
	}
}

// TestVMServerG3StringBuilderEmptyBehavior verifies default/empty argument behavior.
func TestVMServerG3StringBuilderEmptyBehavior(t *testing.T) {
	vm := NewVM(nil, nil, 5)
	host := NewMockHost()
	vm.SetHost(host)

	obj := vm.dispatchNativeCall(nativeObjectServer, "CreateObject", []Value{NewString("g3stringbuilder")})
	if obj.Type != VTNativeObject {
		t.Fatalf("expected VTNativeObject, got %#v", obj)
	}

	_ = vm.dispatchNativeCall(obj.Num, "Append", nil)

	initial := vm.dispatchNativeCall(obj.Num, "ToString", nil)
	if initial.Type != VTString || initial.Str != "" {
		t.Fatalf("expected empty builder result, got %#v", initial)
	}

	vm.dispatchNativeCall(obj.Num, "Append", []Value{NewString("hello")})
	viaPropertyGet := vm.dispatchMemberGet(obj, "ToString")
	if viaPropertyGet.Type != VTString || viaPropertyGet.Str != "hello" {
		t.Fatalf("unexpected property-style ToString result: %#v", viaPropertyGet)
	}
}

// TestVMServerCreateObjectG3Search verifies CreateObject and property dispatch for G3SEARCH.
func TestVMServerCreateObjectG3Search(t *testing.T) {
	resetG3SearchGlobalStateForTests()
	defer resetG3SearchGlobalStateForTests()

	vm := NewVM(nil, nil, 5)
	host := NewMockHost()
	vm.SetHost(host)

	obj := vm.dispatchNativeCall(nativeObjectServer, "CreateObject", []Value{NewString("G3SEARCH")})
	if obj.Type != VTNativeObject {
		t.Fatalf("expected VTNativeObject, got %#v", obj)
	}

	defaultExtension := vm.dispatchMemberGet(obj, "Extension")
	if defaultExtension.Type != VTString || defaultExtension.Str != ".md" {
		t.Fatalf("unexpected default Extension value: %#v", defaultExtension)
	}

	vm.dispatchMemberSet(obj.Num, "IndexPath", NewString("temp/g3search.index"))
	vm.dispatchMemberSet(obj.Num, "DocsPath", NewString("www/manual/md"))
	vm.dispatchMemberSet(obj.Num, "Extension", NewString("txt"))

	indexPath := vm.dispatchMemberGet(obj, "IndexPath")
	docsPath := vm.dispatchMemberGet(obj, "DocsPath")
	extension := vm.dispatchMemberGet(obj, "Extension")

	if indexPath.Type != VTString || indexPath.Str != canonicalIndexPath("temp/g3search.index") {
		t.Fatalf("unexpected IndexPath value: %#v", indexPath)
	}
	if docsPath.Type != VTString || docsPath.Str != canonicalIndexPath("www/manual/md") {
		t.Fatalf("unexpected DocsPath value: %#v", docsPath)
	}
	if extension.Type != VTString || extension.Str != ".txt" {
		t.Fatalf("unexpected Extension value: %#v", extension)
	}
}

// TestVMServerG3SearchBuildAndSearch verifies index build and Search return shape.
func TestVMServerG3SearchBuildAndSearch(t *testing.T) {
	resetG3SearchGlobalStateForTests()
	defer resetG3SearchGlobalStateForTests()

	vm := NewVM(nil, nil, 5)
	host := NewMockHost()
	vm.SetHost(host)

	obj := vm.dispatchNativeCall(nativeObjectServer, "CreateObject", []Value{NewString("G3SEARCH")})
	if obj.Type != VTNativeObject {
		t.Fatalf("expected VTNativeObject, got %#v", obj)
	}

	docsDir := filepath.Join(t.TempDir(), "docs")
	indexDir := filepath.Join(t.TempDir(), "index")
	if err := os.MkdirAll(docsDir, 0o755); err != nil {
		t.Fatalf("mkdir docs dir failed: %v", err)
	}
	if err := os.MkdirAll(indexDir, 0o755); err != nil {
		t.Fatalf("mkdir index dir failed: %v", err)
	}

	content := []byte("AxonASP search validation content")
	filePath := filepath.Join(docsDir, "guide.md")
	if err := os.WriteFile(filePath, content, 0o644); err != nil {
		t.Fatalf("write docs file failed: %v", err)
	}

	vm.dispatchMemberSet(obj.Num, "DocsPath", NewString(docsDir))
	vm.dispatchMemberSet(obj.Num, "IndexPath", NewString(indexDir))
	vm.dispatchMemberSet(obj.Num, "Extension", NewString(".md"))

	_ = vm.dispatchNativeCall(obj.Num, "BuildIndex", nil)
	if vm.lastError != nil {
		t.Fatalf("BuildIndex produced unexpected error: %v", vm.lastError)
	}

	results := vm.dispatchNativeCall(obj.Num, "Search", []Value{NewString("AxonASP")})
	if results.Type != VTArray || results.Arr == nil {
		t.Fatalf("expected VTArray result, got %#v", results)
	}
	if len(results.Arr.Values) == 0 {
		t.Fatalf("expected at least one result row, got %#v", results)
	}

	firstRow := results.Arr.Values[0]
	if firstRow.Type != VTArray || firstRow.Arr == nil {
		t.Fatalf("expected first result row as VTArray, got %#v", firstRow)
	}
	if len(firstRow.Arr.Values) != 2 {
		t.Fatalf("expected [filename, score] tuple, got %#v", firstRow.Arr.Values)
	}

	if firstRow.Arr.Values[0].Type != VTString || firstRow.Arr.Values[0].Str == "" {
		t.Fatalf("expected first tuple item as filename string, got %#v", firstRow.Arr.Values[0])
	}
	if firstRow.Arr.Values[1].Type != VTDouble {
		t.Fatalf("expected second tuple item as score double, got %#v", firstRow.Arr.Values[1])
	}

	if globalBuiltIndexes[canonicalIndexKey(indexDir)] != true {
		t.Fatalf("expected index path to be marked built")
	}
}

// TestVMServerG3SearchBuildIndexWithArgs verifies VB-style BuildIndex(indexPath, docsPath[, ext]).
func TestVMServerG3SearchBuildIndexWithArgs(t *testing.T) {
	resetG3SearchGlobalStateForTests()
	defer resetG3SearchGlobalStateForTests()

	vm := NewVM(nil, nil, 5)
	host := NewMockHost()
	vm.SetHost(host)

	obj := vm.dispatchNativeCall(nativeObjectServer, "CreateObject", []Value{NewString("G3SEARCH")})
	if obj.Type != VTNativeObject {
		t.Fatalf("expected VTNativeObject, got %#v", obj)
	}

	docsDir := filepath.Join(t.TempDir(), "docs")
	indexDir := filepath.Join(t.TempDir(), "index")
	if err := os.MkdirAll(docsDir, 0o755); err != nil {
		t.Fatalf("mkdir docs dir failed: %v", err)
	}
	if err := os.MkdirAll(indexDir, 0o755); err != nil {
		t.Fatalf("mkdir index dir failed: %v", err)
	}

	if err := os.WriteFile(filepath.Join(docsDir, "guide.md"), []byte("linux compatibility guide"), 0o644); err != nil {
		t.Fatalf("write docs file failed: %v", err)
	}

	_ = vm.dispatchNativeCall(obj.Num, "BuildIndex", []Value{NewString(indexDir), NewString(docsDir), NewString("md")})
	if vm.lastError != nil {
		t.Fatalf("BuildIndex with args produced unexpected error: %v", vm.lastError)
	}

	results := vm.dispatchNativeCall(obj.Num, "Search", []Value{NewString("linux")})
	if results.Type != VTArray || results.Arr == nil {
		t.Fatalf("expected VTArray result, got %#v", results)
	}
	if len(results.Arr.Values) == 0 {
		t.Fatalf("expected at least one result row, got %#v", results)
	}
}

// TestVMServerG3SearchSearchSelfHealsStaleEmptyIndex verifies Search auto-rebuilds
// one stale/empty index directory and then returns matches.
func TestVMServerG3SearchSearchSelfHealsStaleEmptyIndex(t *testing.T) {
	resetG3SearchGlobalStateForTests()
	defer resetG3SearchGlobalStateForTests()

	vm := NewVM(nil, nil, 5)
	host := NewMockHost()
	vm.SetHost(host)

	obj := vm.dispatchNativeCall(nativeObjectServer, "CreateObject", []Value{NewString("G3SEARCH")})
	if obj.Type != VTNativeObject {
		t.Fatalf("expected VTNativeObject, got %#v", obj)
	}

	docsDir := filepath.Join(t.TempDir(), "docs")
	indexDir := filepath.Join(t.TempDir(), "index")
	if err := os.MkdirAll(docsDir, 0o755); err != nil {
		t.Fatalf("mkdir docs dir failed: %v", err)
	}
	if err := os.MkdirAll(indexDir, 0o755); err != nil {
		t.Fatalf("mkdir index dir failed: %v", err)
	}

	if err := os.WriteFile(filepath.Join(docsDir, "guide.md"), []byte("AxonASP search validation content"), 0o644); err != nil {
		t.Fatalf("write docs file failed: %v", err)
	}

	// Simulate stale state: marker says built, but no index segments exist.
	if err := os.WriteFile(filepath.Join(indexDir, g3SearchBuiltMarkerFile), []byte("built"), 0o644); err != nil {
		t.Fatalf("write marker file failed: %v", err)
	}
	globalBuiltIndexes[canonicalIndexKey(indexDir)] = true

	vm.dispatchMemberSet(obj.Num, "DocsPath", NewString(docsDir))
	vm.dispatchMemberSet(obj.Num, "IndexPath", NewString(indexDir))
	vm.dispatchMemberSet(obj.Num, "Extension", NewString(".md"))

	results := vm.dispatchNativeCall(obj.Num, "Search", []Value{NewString("AxonASP")})
	if results.Type != VTArray || results.Arr == nil {
		t.Fatalf("expected VTArray result, got %#v", results)
	}
	if len(results.Arr.Values) == 0 {
		t.Fatalf("expected at least one result after self-heal rebuild, got %#v", results)
	}
}

// TestVMServerG3SearchSearchSelfHealsMatchingMarkerEmptyIndex verifies Search
// rebuilds when marker metadata matches docsPath but the index has zero documents.
func TestVMServerG3SearchSearchSelfHealsMatchingMarkerEmptyIndex(t *testing.T) {
	resetG3SearchGlobalStateForTests()
	defer resetG3SearchGlobalStateForTests()

	vm := NewVM(nil, nil, 5)
	host := NewMockHost()
	vm.SetHost(host)

	obj := vm.dispatchNativeCall(nativeObjectServer, "CreateObject", []Value{NewString("G3SEARCH")})
	if obj.Type != VTNativeObject {
		t.Fatalf("expected VTNativeObject, got %#v", obj)
	}

	docsDir := filepath.Join(t.TempDir(), "docs")
	indexDir := filepath.Join(t.TempDir(), "index")
	if err := os.MkdirAll(docsDir, 0o755); err != nil {
		t.Fatalf("mkdir docs dir failed: %v", err)
	}
	if err := os.MkdirAll(indexDir, 0o755); err != nil {
		t.Fatalf("mkdir index dir failed: %v", err)
	}

	if err := os.WriteFile(filepath.Join(docsDir, "guide.md"), []byte("AxonASP linux validation content"), 0o644); err != nil {
		t.Fatalf("write docs file failed: %v", err)
	}

	markerContent := []byte("docs=" + canonicalIndexPath(docsDir))
	if err := os.WriteFile(filepath.Join(indexDir, g3SearchBuiltMarkerFile), markerContent, 0o644); err != nil {
		t.Fatalf("write marker file failed: %v", err)
	}
	globalBuiltIndexes[canonicalIndexKey(indexDir)] = true

	vm.dispatchMemberSet(obj.Num, "DocsPath", NewString(docsDir))
	vm.dispatchMemberSet(obj.Num, "IndexPath", NewString(indexDir))
	vm.dispatchMemberSet(obj.Num, "Extension", NewString(".md"))

	results := vm.dispatchNativeCall(obj.Num, "Search", []Value{NewString("linux")})
	if results.Type != VTArray || results.Arr == nil {
		t.Fatalf("expected VTArray result, got %#v", results)
	}
	if len(results.Arr.Values) == 0 {
		t.Fatalf("expected at least one result after empty-index self-heal, got %#v", results)
	}
}

// TestVMServerG3SearchBuildIndexRunsOncePerPath verifies strict once-per-process behavior.
func TestVMServerG3SearchBuildIndexRunsOncePerPath(t *testing.T) {
	resetG3SearchGlobalStateForTests()
	defer resetG3SearchGlobalStateForTests()

	vm := NewVM(nil, nil, 5)
	host := NewMockHost()
	vm.SetHost(host)

	obj := vm.dispatchNativeCall(nativeObjectServer, "CreateObject", []Value{NewString("G3SEARCH")})
	if obj.Type != VTNativeObject {
		t.Fatalf("expected VTNativeObject, got %#v", obj)
	}

	docsDir := filepath.Join(t.TempDir(), "docs")
	indexDir := filepath.Join(t.TempDir(), "index")
	if err := os.MkdirAll(docsDir, 0o755); err != nil {
		t.Fatalf("mkdir docs dir failed: %v", err)
	}
	if err := os.MkdirAll(indexDir, 0o755); err != nil {
		t.Fatalf("mkdir index dir failed: %v", err)
	}

	fileOne := filepath.Join(docsDir, "one.md")
	if err := os.WriteFile(fileOne, []byte("alpha token"), 0o644); err != nil {
		t.Fatalf("write first docs file failed: %v", err)
	}

	vm.dispatchMemberSet(obj.Num, "DocsPath", NewString(docsDir))
	vm.dispatchMemberSet(obj.Num, "IndexPath", NewString(indexDir))
	vm.dispatchMemberSet(obj.Num, "Extension", NewString(".md"))

	_ = vm.dispatchNativeCall(obj.Num, "BuildIndex", nil)
	if vm.lastError != nil {
		t.Fatalf("first BuildIndex produced unexpected error: %v", vm.lastError)
	}

	fileTwo := filepath.Join(docsDir, "two.md")
	if err := os.WriteFile(fileTwo, []byte("beta token"), 0o644); err != nil {
		t.Fatalf("write second docs file failed: %v", err)
	}

	_ = vm.dispatchNativeCall(obj.Num, "BuildIndex", nil)
	if vm.lastError != nil {
		t.Fatalf("second BuildIndex produced unexpected error: %v", vm.lastError)
	}

	betaResults := vm.dispatchNativeCall(obj.Num, "Search", []Value{NewString("beta")})
	if betaResults.Type != VTArray || betaResults.Arr == nil {
		t.Fatalf("expected VTArray result for beta search, got %#v", betaResults)
	}
	if len(betaResults.Arr.Values) != 0 {
		t.Fatalf("expected no results because second BuildIndex is no-op, got %#v", betaResults.Arr.Values)
	}

	if globalBuiltIndexes[canonicalIndexKey(indexDir)] != true {
		t.Fatalf("expected index path to remain marked built")
	}
}

// TestVMServerG3SearchBuildIndexConcurrentSingleRun ensures only one build runs per path.
func TestVMServerG3SearchBuildIndexConcurrentSingleRun(t *testing.T) {
	resetG3SearchGlobalStateForTests()
	defer resetG3SearchGlobalStateForTests()

	vm := NewVM(nil, nil, 5)
	host := NewMockHost()
	vm.SetHost(host)

	docsDir := filepath.Join(t.TempDir(), "docs")
	indexDir := filepath.Join(t.TempDir(), "index")
	if err := os.MkdirAll(docsDir, 0o755); err != nil {
		t.Fatalf("mkdir docs dir failed: %v", err)
	}
	if err := os.MkdirAll(indexDir, 0o755); err != nil {
		t.Fatalf("mkdir index dir failed: %v", err)
	}

	if err := os.WriteFile(filepath.Join(docsDir, "guide.md"), []byte("single run marker"), 0o644); err != nil {
		t.Fatalf("write docs file failed: %v", err)
	}

	const workers = 12
	var wg sync.WaitGroup
	wg.Add(workers)

	for range workers {
		go func() {
			defer wg.Done()
			search := NewG3Search(vm)
			search.DispatchPropertySet("DocsPath", NewString(docsDir))
			search.DispatchPropertySet("IndexPath", NewString(indexDir))
			search.DispatchPropertySet("Extension", NewString(".md"))
			_ = search.buildIndex()
		}()
	}

	wg.Wait()

	indexKey := canonicalIndexKey(indexDir)
	globalReaderMutex.RLock()
	runs := globalBuildRuns[indexKey]
	built := globalBuiltIndexes[indexKey]
	globalReaderMutex.RUnlock()

	if runs != 1 {
		t.Fatalf("expected exactly one physical build run, got %d", runs)
	}
	if !built {
		t.Fatalf("expected index path to be marked built")
	}

	search := NewG3Search(vm)
	search.DispatchPropertySet("IndexPath", NewString(indexDir))
	results := search.search("single")
	if results.Type != VTArray || results.Arr == nil {
		t.Fatalf("expected VTArray result, got %#v", results)
	}
	if len(results.Arr.Values) == 0 {
		t.Fatalf("expected search results after build, got %#v", results)
	}
}

// TestVMServerG3SearchBuildMarkerPreventsRebuild verifies disk marker dedupes future calls.
func TestVMServerG3SearchBuildMarkerPreventsRebuild(t *testing.T) {
	resetG3SearchGlobalStateForTests()
	defer resetG3SearchGlobalStateForTests()

	vm := NewVM(nil, nil, 5)
	host := NewMockHost()
	vm.SetHost(host)

	docsDir := filepath.Join(t.TempDir(), "docs")
	indexDir := filepath.Join(t.TempDir(), "index")
	if err := os.MkdirAll(docsDir, 0o755); err != nil {
		t.Fatalf("mkdir docs dir failed: %v", err)
	}
	if err := os.MkdirAll(indexDir, 0o755); err != nil {
		t.Fatalf("mkdir index dir failed: %v", err)
	}

	if err := os.WriteFile(filepath.Join(docsDir, "first.md"), []byte("first payload"), 0o644); err != nil {
		t.Fatalf("write first docs file failed: %v", err)
	}

	search := NewG3Search(vm)
	search.DispatchPropertySet("DocsPath", NewString(docsDir))
	search.DispatchPropertySet("IndexPath", NewString(indexDir))
	search.DispatchPropertySet("Extension", NewString(".md"))
	_ = search.buildIndex()

	indexKey := canonicalIndexKey(indexDir)
	globalReaderMutex.RLock()
	firstRuns := globalBuildRuns[indexKey]
	globalReaderMutex.RUnlock()
	if firstRuns != 1 {
		t.Fatalf("expected first build run count to be 1, got %d", firstRuns)
	}

	if !hasPersistentBuildMarker(indexDir) {
		t.Fatalf("expected persistent build marker file")
	}

	resetG3SearchGlobalStateForTests()

	searchAgain := NewG3Search(vm)
	searchAgain.DispatchPropertySet("DocsPath", NewString(docsDir))
	searchAgain.DispatchPropertySet("IndexPath", NewString(indexDir))
	searchAgain.DispatchPropertySet("Extension", NewString(".md"))
	_ = searchAgain.buildIndex()

	globalReaderMutex.RLock()
	secondRuns := globalBuildRuns[indexKey]
	built := globalBuiltIndexes[indexKey]
	globalReaderMutex.RUnlock()
	if secondRuns != 0 {
		t.Fatalf("expected no new build after marker, got %d", secondRuns)
	}
	if !built {
		t.Fatalf("expected path to be marked built from marker")
	}
}

// TestVMServerG3SearchSkipsIndexPathSubtree verifies DocsPath walk skips IndexPath subtree.
func TestVMServerG3SearchSkipsIndexPathSubtree(t *testing.T) {
	resetG3SearchGlobalStateForTests()
	defer resetG3SearchGlobalStateForTests()

	vm := NewVM(nil, nil, 5)
	host := NewMockHost()
	vm.SetHost(host)

	docsDir := filepath.Join(t.TempDir(), "docs")
	indexDir := filepath.Join(docsDir, "index")
	if err := os.MkdirAll(indexDir, 0o755); err != nil {
		t.Fatalf("mkdir index dir failed: %v", err)
	}

	if err := os.WriteFile(filepath.Join(docsDir, "root.md"), []byte("rootterm"), 0o644); err != nil {
		t.Fatalf("write root docs file failed: %v", err)
	}
	if err := os.WriteFile(filepath.Join(indexDir, "nested.md"), []byte("indextoken"), 0o644); err != nil {
		t.Fatalf("write nested docs file failed: %v", err)
	}

	search := NewG3Search(vm)
	search.DispatchPropertySet("DocsPath", NewString(docsDir))
	search.DispatchPropertySet("IndexPath", NewString(indexDir))
	search.DispatchPropertySet("Extension", NewString(".md"))
	_ = search.buildIndex()

	rootResults := search.search("rootterm")
	if rootResults.Type != VTArray || rootResults.Arr == nil || len(rootResults.Arr.Values) == 0 {
		t.Fatalf("expected root term to be indexed")
	}

	nestedResults := search.search("indextoken")
	if nestedResults.Type != VTArray || nestedResults.Arr == nil {
		t.Fatalf("expected VTArray for nested search, got %#v", nestedResults)
	}
	if len(nestedResults.Arr.Values) != 0 {
		t.Fatalf("expected index subtree content to be skipped, got %#v", nestedResults.Arr.Values)
	}
}

// resetG3SearchGlobalStateForTests clears shared G3SEARCH globals between test cases.
func resetG3SearchGlobalStateForTests() {
	globalReaderMutex.Lock()
	_ = closeGlobalReaderLocked()

	for indexPath, waitCh := range globalBuildingIndexes {
		delete(globalBuildingIndexes, indexPath)
		close(waitCh)
	}

	globalBuiltIndexes = make(map[string]bool)
	globalBuildingIndexes = make(map[string]chan struct{})
	globalBuildRuns = make(map[string]int)
	globalReaderMutex.Unlock()
}

// TestVMServerCreateObjectG3Crypto verifies CreateObject aliases for G3Crypto.
func TestVMServerCreateObjectG3Crypto(t *testing.T) {
	vm := NewVM(nil, nil, 5)
	host := NewMockHost()
	vm.SetHost(host)

	aliases := []string{
		"G3CRYPTO",
		"G3.Crypto",
		"System.Security.Cryptography.SHA256CryptoServiceProvider",
	}

	for _, alias := range aliases {
		obj := vm.dispatchNativeCall(nativeObjectServer, "CreateObject", []Value{NewString(alias)})
		if obj.Type != VTNativeObject {
			t.Fatalf("expected VTNativeObject for alias %q, got %#v", alias, obj)
		}
	}
}

// TestVMServerG3CryptoHashing verifies digest methods and hash-related properties.
func TestVMServerG3CryptoHashing(t *testing.T) {
	vm := NewVM(nil, nil, 5)
	host := NewMockHost()
	vm.SetHost(host)

	cryptoObj := vm.dispatchNativeCall(nativeObjectServer, "CreateObject", []Value{NewString("G3Crypto")})
	if cryptoObj.Type != VTNativeObject {
		t.Fatalf("expected VTNativeObject, got %#v", cryptoObj)
	}

	md5Hex := vm.dispatchNativeCall(cryptoObj.Num, "MD5", []Value{NewString("abc")})
	if md5Hex.Type != VTString || md5Hex.Str != "900150983cd24fb0d6963f7d28e17f72" {
		t.Fatalf("unexpected MD5 output: %#v", md5Hex)
	}

	computed := vm.dispatchNativeCall(cryptoObj.Num, "ComputeHash", []Value{NewString("abc"), NewString("sha256")})
	if computed.Type != VTArray || computed.Arr == nil || len(computed.Arr.Values) != 32 {
		t.Fatalf("unexpected ComputeHash output: %#v", computed)
	}

	hashProp := vm.dispatchMemberGet(cryptoObj, "Hash")
	if hashProp.Type != VTArray || hashProp.Arr == nil || len(hashProp.Arr.Values) != 32 {
		t.Fatalf("unexpected Hash property output: %#v", hashProp)
	}

	hashSize := vm.dispatchMemberGet(cryptoObj, "HashSize")
	if hashSize.Type != VTInteger || hashSize.Num != 256 {
		t.Fatalf("unexpected HashSize output: %#v", hashSize)
	}
}

// TestVMServerG3CryptoPasswordFlow verifies bcrypt hash and verify aliases.
func TestVMServerG3CryptoPasswordFlow(t *testing.T) {
	vm := NewVM(nil, nil, 5)
	host := NewMockHost()
	vm.SetHost(host)

	cryptoObj := vm.dispatchNativeCall(nativeObjectServer, "CreateObject", []Value{NewString("G3Crypto")})
	if cryptoObj.Type != VTNativeObject {
		t.Fatalf("expected VTNativeObject, got %#v", cryptoObj)
	}

	hashed := vm.dispatchNativeCall(cryptoObj.Num, "HashPassword", []Value{NewString("Axon#123")})
	if hashed.Type != VTString || hashed.Str == "" {
		t.Fatalf("unexpected HashPassword output: %#v", hashed)
	}

	valid := vm.dispatchNativeCall(cryptoObj.Num, "VerifyPassword", []Value{NewString("Axon#123"), hashed})
	if valid.Type != VTBool || valid.Num != 1 {
		t.Fatalf("expected VerifyPassword=True, got %#v", valid)
	}

	invalid := vm.dispatchNativeCall(cryptoObj.Num, "Verify", []Value{NewString("wrong"), hashed})
	if invalid.Type != VTBool || invalid.Num != 0 {
		t.Fatalf("expected Verify=False, got %#v", invalid)
	}
}

// TestVMServerExecuteUsesChildScopeAndPath verifies Server.Execute keeps child scope isolated and resolves relative paths from the child page.
func TestVMServerExecuteUsesChildScopeAndPath(t *testing.T) {
	rootDir := t.TempDir()
	childDir := filepath.Join(rootDir, "sub")
	if err := os.MkdirAll(childDir, 0o755); err != nil {
		t.Fatalf("mkdir failed: %v", err)
	}
	childPath := filepath.Join(childDir, "child.asp")
	childCode := `<%
Response.Write "child:"
Response.Write Server.MapPath("leaf.asp")
Response.Write ":"
If IsEmpty(parentValue) Then
    Response.Write "empty"
Else
    Response.Write parentValue
End If
%>`
	if err := os.WriteFile(childPath, []byte(childCode), 0o644); err != nil {
		t.Fatalf("write child failed: %v", err)
	}

	parentCode := `<% Dim parentValue : parentValue = "parent" : Response.Write("before|") : Server.Execute("sub/child.asp") : Response.Write("|after:" & parentValue) %>`
	compiler := NewASPCompiler(parentCode)
	compiler.SetSourceName(filepath.Join(rootDir, "parent.asp"))
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVMFromCompiler(compiler)
	host := NewMockHost()
	var out bytes.Buffer
	host.SetOutput(&out)
	host.Server().SetRootDir(rootDir)
	host.Server().SetRequestPath("/parent.asp")
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("run failed: %v", err)
	}
	host.Response().Flush()

	output := strings.ReplaceAll(out.String(), "\\", "/")
	if !strings.Contains(output, "before|child:") {
		t.Fatalf("unexpected output: %q", output)
	}
	if !strings.Contains(output, "/sub/leaf.asp:empty") {
		t.Fatalf("expected child-relative MapPath and isolated scope, got %q", output)
	}
	if !strings.Contains(output, "|after:parent") {
		t.Fatalf("expected parent scope to remain intact, got %q", output)
	}
}

// TestVMRunReturnsScriptTimeout verifies runtime execution aborts when ScriptTimeout is exceeded.
func TestVMRunReturnsScriptTimeout(t *testing.T) {
	compiler := NewASPCompiler(`<% Do : Loop %>`)
	compiler.SetSourceName("timeout.asp")
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVMFromCompiler(compiler)
	host := NewMockHost()
	if err := host.Server().SetScriptTimeout(1); err != nil {
		t.Fatalf("set timeout failed: %v", err)
	}
	vm.SetHost(host)

	err := vm.Run()
	if err == nil {
		t.Fatal("expected script timeout error")
	}
	var axErr *AxonASPError
	if !errors.As(err, &axErr) {
		t.Fatalf("expected AxonASPError, got %T: %v", err, err)
	}
	if axErr.Code != ErrScriptTimeout {
		t.Fatalf("expected ErrScriptTimeout, got %v", axErr.Code)
	}
}

// TestVMRunReturnsResponseBufferLimit verifies Response buffering overflow is surfaced as one AxonASP runtime error.
func TestVMRunReturnsResponseBufferLimit(t *testing.T) {
	compiler := NewASPCompiler(`<%
Response.Buffer = True
Response.Write String(9, "A")
%>`)
	compiler.SetSourceName("buffer.asp")
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVMFromCompiler(compiler)
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	host.Response().SetMaxBufferBytes(8)
	vm.SetHost(host)

	err := vm.Run()
	if err == nil {
		t.Fatal("expected response buffer limit error")
	}
	var axErr *AxonASPError
	if !errors.As(err, &axErr) {
		t.Fatalf("expected AxonASPError, got %T: %v", err, err)
	}
	if axErr.Code != ErrResponseBufferLimitExceeded {
		t.Fatalf("expected ErrResponseBufferLimitExceeded, got %v", axErr.Code)
	}
}

// TestVMServerCreateObjectFSO verifies CreateObject for native Scripting.FileSystemObject.
func TestVMServerCreateObjectFSO(t *testing.T) {
	vm := NewVM(nil, nil, 5)
	host := NewMockHost()
	host.Server().SetRootDir(t.TempDir())
	host.Server().SetRequestPath("/tests/test_server.asp")
	vm.SetHost(host)

	obj := vm.dispatchNativeCall(nativeObjectServer, "CreateObject", []Value{NewString("Scripting.FileSystemObject")})
	if obj.Type != VTNativeObject {
		t.Fatalf("expected VTNativeObject, got %#v", obj)
	}
}

// TestVMServerFSOFileRoundTrip verifies create, write, read, and collection access via FSO.
func TestVMServerFSOFileRoundTrip(t *testing.T) {
	vm := NewVM(nil, nil, 5)
	host := NewMockHost()
	rootDir := t.TempDir()
	host.Server().SetRootDir(rootDir)
	host.Server().SetRequestPath("/tests/test_server.asp")
	vm.SetHost(host)

	fso := vm.dispatchNativeCall(nativeObjectServer, "CreateObject", []Value{NewString("Scripting.FileSystemObject")})
	if fso.Type != VTNativeObject {
		t.Fatalf("expected VTNativeObject for FSO, got %#v", fso)
	}

	vm.dispatchNativeCall(fso.Num, "CreateFolder", []Value{NewString("/sandbox")})

	filePath := "/sandbox/sample.txt"
	textStream := vm.dispatchNativeCall(fso.Num, "CreateTextFile", []Value{NewString(filePath), NewBool(true)})
	if textStream.Type != VTNativeObject {
		t.Fatalf("expected VTNativeObject text stream, got %#v", textStream)
	}

	vm.dispatchNativeCall(textStream.Num, "WriteLine", []Value{NewString("AxonASP")})
	vm.dispatchNativeCall(textStream.Num, "Write", []Value{NewString("FSO")})
	vm.dispatchNativeCall(textStream.Num, "Close", nil)

	exists := vm.dispatchNativeCall(fso.Num, "FileExists", []Value{NewString(filePath)})
	if exists.Type != VTBool || exists.Num != 1 {
		t.Fatalf("expected FileExists True, got %#v", exists)
	}

	reader := vm.dispatchNativeCall(fso.Num, "OpenTextFile", []Value{NewString(filePath), NewInteger(1), NewBool(false)})
	if reader.Type != VTNativeObject {
		t.Fatalf("expected VTNativeObject reader, got %#v", reader)
	}

	content := vm.dispatchNativeCall(reader.Num, "ReadAll", nil)
	if content.Type != VTString || !strings.Contains(content.Str, "AxonASP") || !strings.Contains(content.Str, "FSO") {
		t.Fatalf("unexpected ReadAll content: %#v", content)
	}
	vm.dispatchNativeCall(reader.Num, "Close", nil)

	folder := vm.dispatchNativeCall(fso.Num, "GetFolder", []Value{NewString("/sandbox")})
	if folder.Type != VTNativeObject {
		t.Fatalf("expected VTNativeObject folder, got %#v", folder)
	}

	files := vm.dispatchMemberGet(folder, "Files")
	if files.Type != VTNativeObject {
		t.Fatalf("expected VTNativeObject files collection, got %#v", files)
	}

	count := vm.dispatchMemberGet(files, "Count")
	if count.Type != VTInteger || count.Num != 1 {
		t.Fatalf("unexpected file count: %#v", count)
	}

	item := vm.dispatchNativeCall(files.Num, "Item", []Value{NewString("sample.txt")})
	if item.Type != VTNativeObject {
		t.Fatalf("expected VTNativeObject file item, got %#v", item)
	}

	name := vm.dispatchMemberGet(item, "Name")
	if name.Type != VTString || name.Str != "sample.txt" {
		t.Fatalf("unexpected file name property: %#v", name)
	}

	size := vm.dispatchMemberGet(item, "Size")
	if size.Type != VTInteger || size.Num <= 0 {
		t.Fatalf("unexpected file size property: %#v", size)
	}

	absExpected := filepath.Join(rootDir, "sandbox", "sample.txt")
	pathValue := vm.dispatchMemberGet(item, "Path")
	if pathValue.Type != VTString || !strings.EqualFold(filepath.Clean(pathValue.Str), filepath.Clean(absExpected)) {
		t.Fatalf("unexpected file path property: %#v", pathValue)
	}
}

// TestVMServerFSODrivesAndSpecialFolders verifies cross-platform drive and special folder behavior.
func TestVMServerFSODrivesAndSpecialFolders(t *testing.T) {
	vm := NewVM(nil, nil, 5)
	host := NewMockHost()
	host.Server().SetRootDir(t.TempDir())
	host.Server().SetRequestPath("/tests/test_server.asp")
	vm.SetHost(host)

	fso := vm.dispatchNativeCall(nativeObjectServer, "CreateObject", []Value{NewString("Scripting.FileSystemObject")})
	if fso.Type != VTNativeObject {
		t.Fatalf("expected VTNativeObject for FSO, got %#v", fso)
	}

	drives := vm.dispatchMemberGet(fso, "Drives")
	if drives.Type != VTNativeObject {
		t.Fatalf("expected Drives collection object, got %#v", drives)
	}

	count := vm.dispatchMemberGet(drives, "Count")
	if count.Type != VTInteger || count.Num <= 0 {
		t.Fatalf("expected Drives.Count > 0, got %#v", count)
	}

	tempFolder := vm.dispatchNativeCall(fso.Num, "GetSpecialFolder", []Value{NewInteger(2)})
	if tempFolder.Type != VTString || strings.TrimSpace(tempFolder.Str) == "" {
		t.Fatalf("expected temp folder path, got %#v", tempFolder)
	}
}

// TestVMServerFSOOpenTextFileAppendCreatesFile verifies append mode creates the file when it does not exist.
func TestVMServerFSOOpenTextFileAppendCreatesFile(t *testing.T) {
	vm := NewVM(nil, nil, 5)
	host := NewMockHost()
	rootDir := t.TempDir()
	host.Server().SetRootDir(rootDir)
	host.Server().SetRequestPath("/tests/test_server.asp")
	vm.SetHost(host)

	fso := vm.dispatchNativeCall(nativeObjectServer, "CreateObject", []Value{NewString("Scripting.FileSystemObject")})
	stream := vm.dispatchNativeCall(fso.Num, "OpenTextFile", []Value{NewString("/append/new-file.txt"), NewInteger(8), NewBool(false)})
	if stream.Type != VTNativeObject {
		t.Fatalf("expected text stream object, got %#v", stream)
	}

	vm.dispatchNativeCall(stream.Num, "WriteLine", []Value{NewString("append created")})
	vm.dispatchNativeCall(stream.Num, "Close", nil)

	checkPath := filepath.Join(rootDir, "append", "new-file.txt")
	raw, err := os.ReadFile(checkPath)
	if err != nil {
		t.Fatalf("expected created file, got read error: %v", err)
	}
	if !strings.Contains(string(raw), "append created") {
		t.Fatalf("unexpected append file content: %q", string(raw))
	}
}

// TestVMServerFSONamePropertySetRenames verifies File.Name and Folder.Name
// assignments go through the property-set path used by VBScript member assignment.
func TestVMServerFSONamePropertySetRenames(t *testing.T) {
	vm := NewVM(nil, nil, 5)
	host := NewMockHost()
	rootDir := t.TempDir()
	host.Server().SetRootDir(rootDir)
	host.Server().SetRequestPath("/tests/test_server.asp")
	vm.SetHost(host)

	fso := vm.dispatchNativeCall(nativeObjectServer, "CreateObject", []Value{NewString("Scripting.FileSystemObject")})
	vm.dispatchNativeCall(fso.Num, "CreateFolder", []Value{NewString("/sandbox")})

	stream := vm.dispatchNativeCall(fso.Num, "CreateTextFile", []Value{NewString("/sandbox/original.txt"), NewBool(true)})
	vm.dispatchNativeCall(stream.Num, "Write", []Value{NewString("rename")})
	vm.dispatchNativeCall(stream.Num, "Close", nil)

	fileObj := vm.dispatchNativeCall(fso.Num, "GetFile", []Value{NewString("/sandbox/original.txt")})
	if fileObj.Type != VTNativeObject {
		t.Fatalf("expected file object, got %#v", fileObj)
	}
	vm.dispatchMemberSet(fileObj.Num, "Name", NewString("renamed.txt"))

	if _, err := os.Stat(filepath.Join(rootDir, "sandbox", "original.txt")); !os.IsNotExist(err) {
		t.Fatalf("expected original file to be renamed, stat err=%v", err)
	}
	if _, err := os.Stat(filepath.Join(rootDir, "sandbox", "renamed.txt")); err != nil {
		t.Fatalf("expected renamed file to exist, got %v", err)
	}

	folderObj := vm.dispatchNativeCall(fso.Num, "GetFolder", []Value{NewString("/sandbox")})
	if folderObj.Type != VTNativeObject {
		t.Fatalf("expected folder object, got %#v", folderObj)
	}
	vm.dispatchMemberSet(folderObj.Num, "Name", NewString("renamed-folder"))

	if _, err := os.Stat(filepath.Join(rootDir, "sandbox")); !os.IsNotExist(err) {
		t.Fatalf("expected original folder to be renamed, stat err=%v", err)
	}
	if _, err := os.Stat(filepath.Join(rootDir, "renamed-folder")); err != nil {
		t.Fatalf("expected renamed folder to exist, got %v", err)
	}
}

// TestVMServerFSOMoveFolderReplacesExistingDestination verifies MoveFolder does
// not leave the source folder behind when the destination path already exists.
func TestVMServerFSOMoveFolderReplacesExistingDestination(t *testing.T) {
	vm := NewVM(nil, nil, 5)
	host := NewMockHost()
	rootDir := t.TempDir()
	host.Server().SetRootDir(rootDir)
	host.Server().SetRequestPath("/tests/test_server.asp")
	vm.SetHost(host)

	fso := vm.dispatchNativeCall(nativeObjectServer, "CreateObject", []Value{NewString("Scripting.FileSystemObject")})
	vm.dispatchNativeCall(fso.Num, "CreateFolder", []Value{NewString("/source")})
	vm.dispatchNativeCall(fso.Num, "CreateFolder", []Value{NewString("/dest")})

	stream := vm.dispatchNativeCall(fso.Num, "CreateTextFile", []Value{NewString("/source/keep.txt"), NewBool(true)})
	vm.dispatchNativeCall(stream.Num, "Write", []Value{NewString("source")})
	vm.dispatchNativeCall(stream.Num, "Close", nil)

	stale := vm.dispatchNativeCall(fso.Num, "CreateTextFile", []Value{NewString("/dest/stale.txt"), NewBool(true)})
	vm.dispatchNativeCall(stale.Num, "Write", []Value{NewString("dest")})
	vm.dispatchNativeCall(stale.Num, "Close", nil)

	vm.dispatchNativeCall(fso.Num, "MoveFolder", []Value{NewString("/source"), NewString("/dest")})

	if _, err := os.Stat(filepath.Join(rootDir, "source")); !os.IsNotExist(err) {
		t.Fatalf("expected source folder to be removed after move, stat err=%v", err)
	}
	if _, err := os.Stat(filepath.Join(rootDir, "dest", "keep.txt")); err != nil {
		t.Fatalf("expected moved file at destination, got %v", err)
	}
}

// TestVMServerFSODeleteFolderAbsolutePath verifies DeleteFolder removes a non-empty
// directory when the ASP layer passes a physical absolute path.
func TestVMServerFSODeleteFolderAbsolutePath(t *testing.T) {
	vm := NewVM(nil, nil, 5)
	host := NewMockHost()
	rootDir := t.TempDir()
	host.Server().SetRootDir(rootDir)
	host.Server().SetRequestPath("/tests/test_server.asp")
	vm.SetHost(host)

	fso := vm.dispatchNativeCall(nativeObjectServer, "CreateObject", []Value{NewString("Scripting.FileSystemObject")})
	vm.dispatchNativeCall(fso.Num, "CreateFolder", []Value{NewString("/cleanup")})

	stream := vm.dispatchNativeCall(fso.Num, "CreateTextFile", []Value{NewString("/cleanup/file.txt"), NewBool(true)})
	vm.dispatchNativeCall(stream.Num, "Write", []Value{NewString("cleanup")})
	vm.dispatchNativeCall(stream.Num, "Close", nil)

	fileObj := vm.dispatchNativeCall(fso.Num, "GetFile", []Value{NewString("/cleanup/file.txt")})
	if fileObj.Type != VTNativeObject {
		t.Fatalf("expected cleanup file object, got %#v", fileObj)
	}
	folderObj := vm.dispatchNativeCall(fso.Num, "GetFolder", []Value{NewString("/cleanup")})
	if folderObj.Type != VTNativeObject {
		t.Fatalf("expected cleanup folder object, got %#v", folderObj)
	}
	files := vm.dispatchMemberGet(folderObj, "Files")
	if files.Type != VTNativeObject {
		t.Fatalf("expected cleanup files collection, got %#v", files)
	}

	absPath := filepath.Join(rootDir, "cleanup")
	vm.dispatchNativeCall(fso.Num, "DeleteFolder", []Value{NewString(absPath)})

	if _, err := os.Stat(absPath); !os.IsNotExist(err) {
		t.Fatalf("expected cleanup folder to be deleted, stat err=%v", err)
	}
}

// TestServerExecuteInlinesChildOutput verifies that Server.Execute runs a child ASP file inline,
// sharing the parent's Response so the child output appears in the parent's buffer.
func TestServerExecuteInlinesChildOutput(t *testing.T) {
	tempDir := t.TempDir()

	// Child ASP file that writes a greeting.
	childContent := `<% Response.Write "CHILD" %>`
	childPath := filepath.Join(tempDir, "child.asp")
	if err := os.WriteFile(childPath, []byte(childContent), 0644); err != nil {
		t.Fatalf("write child file: %v", err)
	}

	// Parent ASP uses Server.Execute to include the child inline.
	source := `<%
Response.Write "BEFORE:"
Server.Execute "child.asp"
Response.Write ":AFTER"
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	host.Server().SetRootDir(tempDir)
	host.Server().SetRequestPath("/index.asp")
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	if output.String() != "BEFORE:CHILD:AFTER" {
		t.Fatalf("unexpected Server.Execute output: got %q want %q", output.String(), "BEFORE:CHILD:AFTER")
	}
}

// TestServerExecuteErrorPropagation verifies that Server.Execute propagates errors from a failed script.
func TestServerExecuteErrorPropagation(t *testing.T) {
	tempDir := t.TempDir()

	// Child ASP has division by zero error.
	childContent := `<%
	dim x
	x = 1 / 0
	%>`
	childPath := filepath.Join(tempDir, "child.asp")
	if err := os.WriteFile(childPath, []byte(childContent), 0644); err != nil {
		t.Fatalf("write child file: %v", err)
	}

	// Parent ASP uses Server.Execute and has On Error Resume Next to capture the error details.
	source := `<%
	on error resume next
	Server.Execute "child.asp"
	Response.Write "Err.Number:" & Err.Number & ":Description:" & Err.Description
	%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	host.Server().SetRootDir(tempDir)
	host.Server().SetRequestPath("/index.asp")
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	expectedSub := "Err.Number:-2146828277:Description:Division by zero"
	if !strings.Contains(output.String(), expectedSub) {
		t.Fatalf("unexpected Server.Execute output with division by zero: got %q want to contain %q", output.String(), expectedSub)
	}
}

// TestServerExecuteFileNotFound verifies that Server.Execute on a non-existent file sets the correct ASP error and HRESULT.
func TestServerExecuteFileNotFound(t *testing.T) {
	tempDir := t.TempDir()

	source := `<%
	on error resume next
	Server.Execute "nonexistent.asp"
	Response.Write "Err.Number:" & Err.Number & ":Description:" & Err.Description
	%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	host.Server().SetRootDir(tempDir)
	host.Server().SetRequestPath("/index.asp")
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	// ASP 0173 is path not found / file not found, HRESULT is -2147467259
	expectedSub := "Err.Number:-2147467259:Description:An error occurred when processing the target. The file was not found."
	if !strings.Contains(output.String(), expectedSub) {
		t.Fatalf("unexpected Server.Execute output for nonexistent file: got %q want to contain %q", output.String(), expectedSub)
	}

	// Also verify that without On Error Resume Next, executing a nonexistent file fails with the error returned from Run()
	sourceNoResume := `<%
	Server.Execute "nonexistent.asp"
	%>`
	compiler2 := NewASPCompiler(sourceNoResume)
	if err := compiler2.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm2 := NewVM(compiler2.Bytecode(), compiler2.Constants(), compiler2.GlobalsCount())
	host2 := NewMockHost()
	var output2 bytes.Buffer
	host2.SetOutput(&output2)
	host2.Server().SetRootDir(tempDir)
	host2.Server().SetRequestPath("/index.asp")
	vm2.SetHost(host2)

	err := vm2.Run()
	if err == nil {
		t.Fatalf("expected vm.Run to fail, but it succeeded")
	}

	vmErr, ok := err.(*VMError)
	if !ok {
		t.Fatalf("expected VMError, got %T: %v", err, err)
	}

	if vmErr.ASPCode != 173 || vmErr.Number != -2147467259 {
		t.Fatalf("unexpected error fields: ASPCode=%d, Number=%d", vmErr.ASPCode, vmErr.Number)
	}
}

// TestServerTransferStopsParentAndRunsChild verifies that Server.Transfer runs the child file but

// stops the parent script from producing any further output.
func TestServerTransferStopsParentAndRunsChild(t *testing.T) {
	tempDir := t.TempDir()

	// Child ASP writes its own output.
	childContent := `<% Response.Write "TRANSFERRED" %>`
	childPath := filepath.Join(tempDir, "child.asp")
	if err := os.WriteFile(childPath, []byte(childContent), 0644); err != nil {
		t.Fatalf("write child file: %v", err)
	}

	// Parent writes something, then transfers (parent output should be cleared, child executes).
	source := `<%
Response.Write "THIS SHOULD BE CLEARED"
Server.Transfer "child.asp"
Response.Write "THIS SHOULD NOT APPEAR"
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	host.Server().SetRootDir(tempDir)
	host.Server().SetRequestPath("/index.asp")
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	if output.String() != "TRANSFERRED" {
		t.Fatalf("unexpected Server.Transfer output: got %q want %q", output.String(), "TRANSFERRED")
	}
}

// TestObjectContextSetCompleteFiresHandler verifies that ObjectContext.SetComplete fires
// OnTransactionCommit after the script ends.
func TestObjectContextSetCompleteFiresHandler(t *testing.T) {
	source := `<%
Sub OnTransactionCommit()
    Response.Write "COMMITTED"
End Sub
ObjectContext.SetComplete
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

	if output.String() != "COMMITTED" {
		t.Fatalf("unexpected ObjectContext.SetComplete output: got %q want %q", output.String(), "COMMITTED")
	}
}

// TestObjectContextSetAbortFiresHandler verifies that ObjectContext.SetAbort fires
// OnTransactionAbort after the script ends.
func TestObjectContextSetAbortFiresHandler(t *testing.T) {
	source := `<%
Sub OnTransactionAbort()
    Response.Write "ABORTED"
End Sub
ObjectContext.SetAbort
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

	if output.String() != "ABORTED" {
		t.Fatalf("unexpected ObjectContext.SetAbort output: got %q want %q", output.String(), "ABORTED")
	}
}

// TestServerExecuteResponseRedirectPropagation verifies that Response.Redirect inside
// a Server.Execute'd child file terminates the ENTIRE page, not just the child scope.
// This matches IIS Classic ASP behavior.
func TestServerExecuteResponseRedirectPropagation(t *testing.T) {
	tempDir := t.TempDir()

	// Child ASP file that redirects.
	childContent := `<% Response.Redirect "target.asp" %>`
	childPath := filepath.Join(tempDir, "child.asp")
	if err := os.WriteFile(childPath, []byte(childContent), 0644); err != nil {
		t.Fatalf("write child file: %v", err)
	}

	// Parent ASP uses Server.Execute — any output AFTER Server.Execute must NOT appear
	// because Response.Redirect inside the child terminates the entire page.
	source := `<%
Dim afterFlag
afterFlag = False
Response.Write "before|"
Server.Execute "child.asp"
afterFlag = True
Response.Write "|after"
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	afterFlagIdx, afterFlagExists := compiler.Globals.Get("afterFlag")
	if !afterFlagExists {
		t.Fatal("afterFlag global not found")
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	host.Server().SetRootDir(tempDir)
	host.Server().SetRequestPath("/index.asp")
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	// afterFlag must remain False because Response.Redirect should stop the entire page.
	if vm.Globals[afterFlagIdx].Type == VTBool && vm.Globals[afterFlagIdx].Num != 0 {
		t.Fatal("BUG: Server.Execute did NOT propagate Response.Redirect — afterFlag was set to True")
	}

	// Redirect status must be set.
	if status := host.Response().GetStatus(); status != "302 Found" {
		t.Fatalf("expected 302 Found status after Response.Redirect in Server.Execute, got %q", status)
	}

	// Output must be empty because Response.Redirect clears the entire buffer
	// (including what the parent wrote before Server.Execute).
	if output.String() != "" {
		t.Fatalf("unexpected output: %q (expected empty because Response.Redirect clears the buffer)", output.String())
	}
}

// TestServerExecuteResponseEndPropagation verifies that Response.End inside
// a Server.Execute'd child file terminates the ENTIRE page.
func TestServerExecuteResponseEndPropagation(t *testing.T) {
	tempDir := t.TempDir()

	// Child ASP file that ends the response.
	childContent := `<% Response.End %>`
	childPath := filepath.Join(tempDir, "child.asp")
	if err := os.WriteFile(childPath, []byte(childContent), 0644); err != nil {
		t.Fatalf("write child file: %v", err)
	}

	// Parent ASP uses Server.Execute — anything after must NOT appear.
	source := `<%
Dim afterFlag
afterFlag = False
Response.Write "before|"
Server.Execute "child.asp"
afterFlag = True
Response.Write "|after"
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	afterFlagIdx, afterFlagExists := compiler.Globals.Get("afterFlag")
	if !afterFlagExists {
		t.Fatal("afterFlag global not found")
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	host.Server().SetRootDir(tempDir)
	host.Server().SetRequestPath("/index.asp")
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	// afterFlag must remain False.
	if vm.Globals[afterFlagIdx].Type == VTBool && vm.Globals[afterFlagIdx].Num != 0 {
		t.Fatal("BUG: Server.Execute did NOT propagate Response.End — afterFlag was set to True")
	}

	// Response.End flushes the buffer before terminating, so "before|" should appear.
	// But "|after" must NOT appear because execution was terminated.
	if !strings.Contains(output.String(), "before|") {
		t.Fatalf("expected output to contain 'before|', got %q", output.String())
	}
	if strings.Contains(output.String(), "|after") {
		t.Fatalf("output must NOT contain '|after' because Response.End should terminate the page, got %q", output.String())
	}
}

// TestServerExecuteResponseRedirectBareCallInChild verifies that a bare Sub call
// WITH arguments inside a Server.Execute'd child file, where the Sub calls
// Response.Redirect, terminates the ENTIRE page (both child and parent).
func TestServerExecuteResponseRedirectBareCallInChild(t *testing.T) {
	tempDir := t.TempDir()

	// Child ASP defines a Sub that does Response.Redirect and calls it with a bare call.
	childContent := `<%
Sub Inner(target, msg)
    Response.Redirect "target.asp?act=" & target & "&msg=" & msg
End Sub
Inner "index", "success"
%>`
	childPath := filepath.Join(tempDir, "child.asp")
	if err := os.WriteFile(childPath, []byte(childContent), 0644); err != nil {
		t.Fatalf("write child file: %v", err)
	}

	// Parent uses Server.Execute to run the child.
	source := `<%
Dim afterFlag
afterFlag = False
Response.Write "before|"
Server.Execute "child.asp"
afterFlag = True
Response.Write "|after"
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	afterFlagIdx, afterFlagExists := compiler.Globals.Get("afterFlag")
	if !afterFlagExists {
		t.Fatal("afterFlag global not found")
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	host.Server().SetRootDir(tempDir)
	host.Server().SetRequestPath("/index.asp")
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	// afterFlag must remain False.
	if vm.Globals[afterFlagIdx].Type == VTBool && vm.Globals[afterFlagIdx].Num != 0 {
		t.Fatal("BUG: Server.Execute with bare call child did NOT propagate Response.Redirect — afterFlag was set to True")
	}

	if status := host.Response().GetStatus(); status != "302 Found" {
		t.Fatalf("expected 302 Found status after Response.Redirect in bare call child, got %q", status)
	}

	// Output must be empty because Response.Redirect clears the buffer.
	if output.String() != "" {
		t.Fatalf("unexpected output: %q (expected empty because Response.Redirect clears the buffer)", output.String())
	}
}
