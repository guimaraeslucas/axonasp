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

// TestVMRequestMemberChainExpression verifies expression-chain behavior for Request collections.
func TestVMRequestMemberChainExpression(t *testing.T) {
	source := `<%= Request.QueryString.Count %>|<%= Request.QueryString.Key(1) %>|<%= Request.QueryString("a") %>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	host.Request().QueryString.Add("a", "1")
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}

	if output.String() != "1|a|1" {
		t.Fatalf("unexpected output: %q", output.String())
	}
}

// TestVMRequestBinaryReadExpression verifies Request.BinaryRead in expression call path.
func TestVMRequestBinaryReadExpression(t *testing.T) {
	source := `<%= Request.TotalBytes %>|<%= Request.BinaryRead(3) %>|<%= Request.BinaryRead(10) %>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	host.Request().SetBody([]byte("abcdef"))
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}

	if output.String() != "6|abc|def" {
		t.Fatalf("unexpected output: %q", output.String())
	}
}

// TestVMRequestBinaryReadByRefCountUpdate verifies Request.BinaryRead updates the ByRef count argument.
func TestVMRequestBinaryReadByRefCountUpdate(t *testing.T) {
	source := `<%
Dim readBytes, payload, loops, tmp
readBytes = 200000
payload = Request.BinaryRead(readBytes)
Do Until readBytes < 1
	tmp = Request.BinaryRead(readBytes)
	If readBytes > 0 Then
		payload = payload & MidB(tmp, 1, LenB(tmp))
	End If
	loops = loops + 1
	If loops > 5 Then Exit Do
Loop
Response.Write LenB(payload) & "|" & readBytes
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	host.Request().SetBody([]byte("abcdef"))
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	if output.String() != "6|0" {
		t.Fatalf("expected uploader-compatible read loop output 6|0, got %q", output.String())
	}
}

// TestVMRequestBinaryReadMidBPreservesBinaryLength verifies that byte-oriented
// operations do not shrink BinaryRead payloads containing UTF-8-valid byte runs.
func TestVMRequestBinaryReadMidBPreservesBinaryLength(t *testing.T) {
	source := `<%
Dim n, b, p
n = Request.TotalBytes
b = Request.BinaryRead(n)
p = MidB(b, 1, LenB(b))
Response.Write LenB(p)
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	host.Request().SetBody([]byte{0xE2, 0x82, 0xAC, 0xC3, 0xA9, 0x41, 0xFF})
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	if output.String() != "7" {
		t.Fatalf("expected preserved binary length 7, got %q", output.String())
	}
}

// TestVMRequestBinaryReadMultipartUploaderCompatibility verifies that uploader-style
// ADODB.Stream binary/text round-trips preserve the multipart payload bytes exactly.
func TestVMRequestBinaryReadMultipartUploaderCompatibility(t *testing.T) {
	rootDir := t.TempDir()
	if err := os.MkdirAll(filepath.Join(rootDir, "sandbox"), 0755); err != nil {
		t.Fatalf("failed to create sandbox directory: %v", err)
	}
	body := []byte("------AxonBoundary\r\n" +
		"Content-Disposition: form-data; name=\"upload\"; filename=\"sample.bin\"\r\n" +
		"Content-Type: application/octet-stream\r\n\r\n")
	body = append(body, []byte{0x41, 0x00, 0xFF, 0x42, 0x43, 0x7F}...)
	body = append(body, []byte("\r\n------AxonBoundary--\r\n")...)

	source := `<%
Dim readBytes, bodyData, tmpBinRequest
Dim inStream, outStream, bodyText

Set inStream = Server.CreateObject("ADODB.Stream")
inStream.Type = 1
inStream.Open

readBytes = 200000
bodyData = Request.BinaryRead(readBytes)
bodyData = MidB(bodyData, 1, LenB(bodyData))
Do Until readBytes < 1
	tmpBinRequest = Request.BinaryRead(readBytes)
	If readBytes > 0 Then
		bodyData = bodyData & MidB(tmpBinRequest, 1, LenB(tmpBinRequest))
	End If
Loop

inStream.Write bodyData
inStream.Position = 0
inStream.Type = 2
inStream.Charset = "iso-8859-1"
bodyText = inStream.ReadText

Set outStream = Server.CreateObject("ADODB.Stream")
outStream.Type = 2
outStream.Charset = "iso-8859-1"
outStream.Open
outStream.WriteText bodyText
outStream.Position = 0
outStream.Type = 1
outStream.SaveToFile Server.MapPath("/sandbox/out.bin"), 2
outStream.Close
inStream.Close

Response.Write Len(bodyText)
%>`

	compiler := NewASPCompiler(source)
	compiler.SetSourceName("www/tests/uploader_vm_regression.asp")
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	host.Server().SetRootDir(rootDir)
	host.Server().SetRequestPath("/tests/uploader_vm_regression.asp")
	host.Request().SetBody(body)
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	savedPath := filepath.Join(rootDir, "sandbox", "out.bin")
	saved, err := os.ReadFile(savedPath)
	if err != nil {
		t.Fatalf("failed to read saved upload: %v", err)
	}
	want := body
	if !bytes.Equal(saved, want) {
		t.Fatalf("unexpected saved upload bytes: output=%q got %v want %v", output.String(), saved, want)
	}
}

// TestVMRequestBinaryReadAspLiteUploaderFileStart verifies the exact aspLite
// file extraction path keeps the first uploaded byte.
func TestVMRequestBinaryReadAspLiteUploaderFileStart(t *testing.T) {
	rootDir := t.TempDir()
	if err := os.MkdirAll(filepath.Join(rootDir, "sandbox"), 0755); err != nil {
		t.Fatalf("failed to create sandbox directory: %v", err)
	}
	filePayload := []byte("%PDF-1.4\r\n")
	body := []byte("------AxonBoundary\r\n" +
		"Content-Disposition: form-data; name=\"upload\"; filename=\"sample.pdf\"\r\n" +
		"Content-Type: application/pdf\r\n\r\n")
	body = append(body, filePayload...)
	body = append(body, []byte("\r\n------AxonBoundary--\r\n")...)

	source := `<%
Function String2Byte(sString)
	Dim i
	For i = 1 To Len(sString)
		String2Byte = String2Byte & ChrB(AscB(Mid(sString, i, 1)))
	Next
End Function

Function FindToken(sToken, nStart, data)
	FindToken = InStrB(nStart, data, sToken)
End Function

Function SkipToken(sToken, nStart, data)
	SkipToken = InStrB(nStart, data, sToken)
	SkipToken = SkipToken + LenB(sToken)
End Function

Function ExtractField(sToken, nStart, data)
	Dim nEnd, fieldBytes, j
	nEnd = InStrB(nStart, data, sToken)
	fieldBytes = MidB(data, nStart, nEnd - nStart)
	For j = 1 To LenB(fieldBytes)
		ExtractField = ExtractField & Chr(AscB(MidB(fieldBytes, j, 1)))
	Next
End Function

Dim StreamRequest, VarArrayBinRequest, tmpBinRequest, readBytes
Dim vDataSep, tNewLine, tDoubleQuotes, tTerm, tFilename, tName, tContentDisp, tContentType
Dim nCurPos, nDataBoundPos, nLastSepPos, nPosFile, nPosBound, auxStr
Dim fileStart, fileLength, streamFile

Set StreamRequest = Server.CreateObject("ADODB.Stream")
StreamRequest.Type = 2
StreamRequest.Open

tNewLine = String2Byte(Chr(13))
tDoubleQuotes = String2Byte(Chr(34))
tTerm = String2Byte("--")
tFilename = String2Byte("filename=""")
tName = String2Byte("name=""")
tContentDisp = String2Byte("Content-Disposition")
tContentType = String2Byte("Content-Type:")

readBytes = 200000
VarArrayBinRequest = Request.BinaryRead(readBytes)
VarArrayBinRequest = MidB(VarArrayBinRequest, 1, LenB(VarArrayBinRequest))
Do Until readBytes < 1
	tmpBinRequest = Request.BinaryRead(readBytes)
	If readBytes > 0 Then
		VarArrayBinRequest = VarArrayBinRequest & MidB(tmpBinRequest, 1, LenB(tmpBinRequest))
	End If
Loop
StreamRequest.WriteText VarArrayBinRequest
StreamRequest.Flush

nCurPos = FindToken(tNewLine, 1, VarArrayBinRequest)
vDataSep = MidB(VarArrayBinRequest, 1, nCurPos - 1)
nDataBoundPos = 1
nLastSepPos = FindToken(vDataSep & tTerm, 1, VarArrayBinRequest)
nCurPos = SkipToken(tContentDisp, nDataBoundPos, VarArrayBinRequest)
nCurPos = SkipToken(tName, nCurPos, VarArrayBinRequest)
auxStr = ExtractField(tDoubleQuotes, nCurPos, VarArrayBinRequest)
nPosFile = FindToken(tFilename, nCurPos, VarArrayBinRequest)
nPosBound = FindToken(vDataSep, nCurPos, VarArrayBinRequest)
nCurPos = SkipToken(tFilename, nCurPos, VarArrayBinRequest)
auxStr = ExtractField(tDoubleQuotes, nCurPos, VarArrayBinRequest)
nCurPos = SkipToken(tContentType, nCurPos, VarArrayBinRequest)
auxStr = ExtractField(tNewLine, nCurPos, VarArrayBinRequest)
nCurPos = FindToken(tNewLine, nCurPos, VarArrayBinRequest) + 4
fileStart = nCurPos + 1
fileLength = FindToken(vDataSep, nCurPos, VarArrayBinRequest) - 2 - nCurPos

Set streamFile = Server.CreateObject("ADODB.Stream")
streamFile.Type = 1
streamFile.Open
StreamRequest.Position = fileStart
StreamRequest.CopyTo streamFile, fileLength
streamFile.SaveToFile Server.MapPath("/sandbox/sample.pdf"), 2
streamFile.Close
StreamRequest.Close

Response.Write fileStart & "|" & fileLength
%>`

	compiler := NewASPCompiler(source)
	compiler.SetSourceName("www/tests/uploader_file_start_regression.asp")
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	host.Server().SetRootDir(rootDir)
	host.Server().SetRequestPath("/tests/uploader_file_start_regression.asp")
	host.Request().SetBody(body)
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	savedPath := filepath.Join(rootDir, "sandbox", "sample.pdf")
	saved, err := os.ReadFile(savedPath)
	if err != nil {
		t.Fatalf("failed to read saved file: %v", err)
	}
	if !bytes.Equal(saved, filePayload) {
		t.Fatalf("unexpected uploader extraction bytes: output=%q got=%q want=%q", output.String(), string(saved), string(filePayload))
	}
}

// TestForEachRequestForm verifies that For Each over Request.Form yields form field names.
func TestForEachRequestForm(t *testing.T) {
	source := `<%
Dim k
For Each k In Request.Form
    Response.Write k & "|"
Next
%>`
	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}
	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	host.Request().Form.Add("username", "alice")
	host.Request().Form.Add("email", "alice@test.com")
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	if output.String() != "username|email|" {
		t.Fatalf("expected username|email|, got %q", output.String())
	}
}

// TestForEachRequestQueryString verifies that For Each over Request.QueryString yields field names.
func TestForEachRequestQueryString(t *testing.T) {
	source := `<%
Dim k
For Each k In Request.QueryString
    Response.Write k & "|"
Next
%>`
	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}
	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	host.Request().QueryString.Add("page", "1")
	host.Request().QueryString.Add("sort", "asc")
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	if output.String() != "page|sort|" {
		t.Fatalf("expected page|sort|, got %q", output.String())
	}
}
