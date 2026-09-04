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
	"fmt"
	"net"
	"net/http"
	"net/http/httptest"
	"path/filepath"
	"strings"
	"syscall"
	"testing"
)

// runASPSource compiles and executes one inline ASP program and returns the rendered output.
func runASPSource(t *testing.T, source string, configureHost func(*MockHost)) string {
	t.Helper()
	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}
	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	if configureHost != nil {
		configureHost(host)
	}
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()
	return output.String()
}

func TestMSXML(t *testing.T) {
	vm := NewVM(nil, nil, 0)
	vm.host = &MockHost{}

	// Test MSXML2.DOMDocument
	dom := NewMsXML2DOMDocument(vm)
	if dom == nil {
		t.Fatal("Failed to create DOMDocument")
	}

	res := dom.DispatchMethod("loadXML", []Value{NewString("<root><item>Hello</item></root>")})
	if res.Type != VTBool || res.Num == 0 {
		t.Fatalf("loadXML failed")
	}

	nodes := dom.DispatchMethod("selectNodes", []Value{NewString("//item")})
	if nodes.Type != VTNativeObject {
		t.Fatalf("expected node list, got %v", nodes.Type)
	}

	nodeList := vm.msxmlNodeListItems[nodes.Num]
	if nodeList == nil {
		t.Fatalf("NodeList missing")
	}

	length := nodeList.DispatchPropertyGet("length")
	if length.Num != 1 {
		t.Errorf("expected 1 item, got %d", length.Num)
	}

	// Test MSXML2.ServerXMLHTTP (just structure)
	httpLib := NewMsXML2ServerXMLHTTP(vm)
	if httpLib == nil {
		t.Fatal("Failed to create ServerXMLHTTP")
	}
}

// TestForEachMSXMLNodeList verifies that For Each over xml.selectNodes() yields XML element nodes.
func TestForEachMSXMLNodeList(t *testing.T) {
	source := `<%
Dim xml, nodes, node, out
Set xml = CreateObject("MSXML2.DOMDocument")
xml.loadXML("<root><item>alpha</item><item>beta</item></root>")
Set nodes = xml.selectNodes("//item")
For Each node In nodes
    out = out & node.text & "|"
Next
Response.Write out
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

	if output.String() != "alpha|beta|" {
		t.Fatalf("expected alpha|beta|, got %q", output.String())
	}
}

func TestMSXMLLoadXMLPopulatesParseErrorDetails(t *testing.T) {
	vm := NewVM(nil, nil, 0)
	vm.host = &MockHost{}

	dom := NewMsXML2DOMDocument(vm)
	if dom == nil {
		t.Fatal("Failed to create DOMDocument")
	}

	res := dom.DispatchMethod("loadXML", []Value{NewString("<root><a></root")})
	if res.Type != VTBool || res.Num != 0 {
		t.Fatalf("expected invalid xml loadXML result False, got %#v", res)
	}

	parseErrorVal := dom.DispatchPropertyGet("parseError")
	if parseErrorVal.Type != VTNativeObject {
		t.Fatalf("expected parseError native object, got %#v", parseErrorVal)
	}

	parseErrorObj := vm.msxmlParseErrorItems[parseErrorVal.Num]
	if parseErrorObj == nil {
		t.Fatalf("expected parseError object in vm table for id %d", parseErrorVal.Num)
	}

	if parseErrorObj.DispatchPropertyGet("reason").String() == "" {
		t.Fatalf("expected parseError.reason to be populated")
	}
	if parseErrorObj.DispatchPropertyGet("line").Num <= 0 {
		t.Fatalf("expected parseError.line > 0, got %d", parseErrorObj.DispatchPropertyGet("line").Num)
	}
	if parseErrorObj.DispatchPropertyGet("linepos").Num <= 0 {
		t.Fatalf("expected parseError.linepos > 0, got %d", parseErrorObj.DispatchPropertyGet("linepos").Num)
	}
}

func TestMSXMLDOMDocumentSetPropertyMethod(t *testing.T) {
	vm := NewVM(nil, nil, 0)
	vm.host = &MockHost{}

	dom := NewMsXML2DOMDocument(vm)
	if dom == nil {
		t.Fatal("Failed to create DOMDocument")
	}

	res := dom.DispatchMethod("setProperty", []Value{NewString("ServerHTTPRequest"), NewBool(true)})
	if res.Type != VTEmpty {
		t.Fatalf("expected empty return from setProperty, got %#v", res)
	}

	serverHTTPRequest := dom.DispatchPropertyGet("ServerHTTPRequest")
	if serverHTTPRequest.Type != VTBool || serverHTTPRequest.Num == 0 {
		t.Fatalf("expected ServerHTTPRequest=True after setProperty, got %#v", serverHTTPRequest)
	}
}

func TestMSXMLNodeListDefaultItemAccess(t *testing.T) {
	output := runASPSource(t, `<%
Dim xml, nodes, node
Set xml = CreateObject("MSXML2.DOMDocument")
xml.loadXML("<root><item>alpha</item><item>beta</item></root>")
Set nodes = xml.getElementsByTagName("item")
Set node = nodes(0)
Response.Write node.nodeName & "|" & node.text
%>`, nil)

	if output != "item|alpha" {
		t.Fatalf("expected item|alpha, got %q", output)
	}
}

// TestMSXMLLoadXMLRejectsMismatchedClosingTags verifies malformed XML returns False and populates ParseError.
func TestMSXMLLoadXMLRejectsMismatchedClosingTags(t *testing.T) {
	vm := NewVM(nil, nil, 0)
	vm.host = &MockHost{}

	dom := NewMsXML2DOMDocument(vm)
	res := dom.DispatchMethod("loadXML", []Value{NewString("<root><broken></root>")})
	if res.Type != VTBool || res.Num != 0 {
		t.Fatalf("expected mismatched-tag loadXML result False, got %#v", res)
	}

	parseErrorVal := dom.DispatchPropertyGet("parseError")
	parseErrorObj := vm.msxmlParseErrorItems[parseErrorVal.Num]
	if parseErrorObj == nil {
		t.Fatal("expected parseError object")
	}
	if parseErrorObj.ErrorReason == "" {
		t.Fatal("expected parseError reason to be populated")
	}
	if parseErrorObj.Line <= 0 || parseErrorObj.LinePos <= 0 {
		t.Fatalf("expected positive parseError position, got line=%d linepos=%d", parseErrorObj.Line, parseErrorObj.LinePos)
	}
	if parseErrorObj.SrcText != "<root><broken></root>" {
		t.Fatalf("expected parseError srcText to preserve input, got %q", parseErrorObj.SrcText)
	}
}

// TestMSXMLNodeListNextNodeIterates verifies nextNode returns each selected node once before exhaustion.
func TestMSXMLNodeListNextNodeIterates(t *testing.T) {
	vm := NewVM(nil, nil, 0)
	vm.host = &MockHost{}

	dom := NewMsXML2DOMDocument(vm)
	res := dom.DispatchMethod("loadXML", []Value{NewString("<root><item>alpha</item><item>beta</item></root>")})
	if res.Type != VTBool || res.Num == 0 {
		t.Fatalf("loadXML failed: %#v", res)
	}
	nodesVal := dom.DispatchMethod("selectNodes", []Value{NewString("//item")})
	nodeList := vm.msxmlNodeListItems[nodesVal.Num]
	if nodeList == nil {
		t.Fatal("expected node list")
	}
	first := nodeList.DispatchMethod("nextNode", nil)
	second := nodeList.DispatchMethod("nextNode", nil)
	third := nodeList.DispatchMethod("nextNode", nil)
	if first.Type != VTNativeObject || second.Type != VTNativeObject {
		t.Fatalf("expected first two nextNode calls to return native objects, got %#v %#v", first, second)
	}
	if third.Type != VTEmpty {
		t.Fatalf("expected third nextNode call to return Empty, got %#v", third)
	}
	if got := vm.msxmlElementItems[first.Num].DispatchPropertyGet("text").String(); got != "alpha" {
		t.Fatalf("expected first nextNode text alpha, got %q", got)
	}
	if got := vm.msxmlElementItems[second.Num].DispatchPropertyGet("text").String(); got != "beta" {
		t.Fatalf("expected second nextNode text beta, got %q", got)
	}
}

// TestMSXMLPreserveWhiteSpaceRetainsTextNodes verifies whitespace-only nodes survive when PreserveWhiteSpace=True.
func TestMSXMLPreserveWhiteSpaceRetainsTextNodes(t *testing.T) {
	vm := NewVM(nil, nil, 0)
	vm.host = &MockHost{}

	dom := NewMsXML2DOMDocument(vm)
	dom.DispatchPropertySet("preserveWhiteSpace", []Value{NewBool(true)})
	res := dom.DispatchMethod("loadXML", []Value{NewString("<root>  <a>1</a>  <b>2</b> </root>")})
	if res.Type != VTBool || res.Num == 0 {
		t.Fatalf("expected preserveWhiteSpace loadXML success, got %#v", res)
	}
	rootVal := dom.DispatchPropertyGet("documentElement")
	root := vm.msxmlElementItems[rootVal.Num]
	if root == nil {
		t.Fatal("expected root element")
	}
	firstChild := root.DispatchPropertyGet("firstChild")
	if firstChild.Type != VTNativeObject {
		t.Fatalf("expected firstChild native object, got %#v", firstChild)
	}
	if got := vm.msxmlElementItems[firstChild.Num].Name; got != "#text" {
		t.Fatalf("expected firstChild to be whitespace text node, got %q", got)
	}
}

// TestMSXMLAppendChildFromASP verifies ASP can build one DOM tree using AppendChild on created elements.
func TestMSXMLAppendChildFromASP(t *testing.T) {
	output := runASPSource(t, `<%
Dim doc, root, child, textNode
Set doc = CreateObject("MSXML2.DOMDocument")
Set root = doc.createElement("root")
Set child = doc.createElement("child")
Set textNode = doc.createTextNode("hello")
Call child.appendChild(textNode)
Call root.appendChild(child)
Call doc.appendChild(root)
Response.Write doc.documentElement.nodeName & "|" & doc.selectSingleNode("//child").text
%>`, nil)

	if output != "root|hello" {
		t.Fatalf("expected root|hello, got %q", output)
	}
}

// TestMSXMLSaveLoadRoundTripKeepsTextSingle verifies save/load roundtrips do not duplicate element text.
func TestMSXMLSaveLoadRoundTripKeepsTextSingle(t *testing.T) {
	vm := NewVM(nil, nil, 0)
	host := NewMockHost()
	rootDir := t.TempDir()
	host.Server().SetRootDir(rootDir)
	host.Server().SetRequestPath("/page.asp")
	vm.SetHost(host)

	dom := NewMsXML2DOMDocument(vm)
	if res := dom.DispatchMethod("loadXML", []Value{NewString("<root><item><name>Gamma Guide</name></item></root>")}); res.Type != VTBool || res.Num == 0 {
		t.Fatalf("loadXML failed: %#v", res)
	}
	if res := dom.DispatchMethod("save", []Value{NewString("roundtrip.xml")}); res.Type != VTBool || res.Num == 0 {
		t.Fatalf("save failed: %#v", res)
	}

	reloaded := NewMsXML2DOMDocument(vm)
	if res := reloaded.DispatchMethod("load", []Value{NewString("roundtrip.xml")}); res.Type != VTBool || res.Num == 0 {
		t.Fatalf("load failed: %#v", res)
	}
	nodeVal := reloaded.DispatchMethod("selectSingleNode", []Value{NewString("//name")})
	if nodeVal.Type != VTNativeObject {
		t.Fatalf("expected selectSingleNode native object, got %#v", nodeVal)
	}
	got := vm.msxmlElementItems[nodeVal.Num].DispatchPropertyGet("text").String()
	if got != "Gamma Guide" {
		t.Fatalf("expected single text after roundtrip, got %q (file %s)", got, filepath.Join(rootDir, "roundtrip.xml"))
	}
}

type mockRoundTripper struct {
	roundTripFunc func(req *http.Request) (*http.Response, error)
}

func (m *mockRoundTripper) RoundTrip(req *http.Request) (*http.Response, error) {
	return m.roundTripFunc(req)
}

// TestMSXMLServerXMLHTTPOpenRejectsRelativeAndUnrecognizedScheme verifies that Open with relative
// or unrecognized protocol schemes raises COM error 0x80072EE6 (-2147012890).
func TestMSXMLServerXMLHTTPOpenRejectsRelativeAndUnrecognizedScheme(t *testing.T) {
	vm := NewVM(nil, nil, 0)
	vm.host = &MockHost{}

	testCases := []string{
		"/relative/path.asp",
		"relative/path.asp",
		"ftp://example.com/file.txt",
		"file:///C:/test.txt",
		"ccte://custom/protocol",
		"",
	}

	for _, tc := range testCases {
		t.Run(tc, func(t *testing.T) {
			httpLib := NewMsXML2ServerXMLHTTP(vm)
			_, err := httpLib.legacyCallMethod("open", "GET", tc)
			if err == nil {
				t.Fatalf("expected open to fail for url %q", tc)
			}
			vme, ok := err.(*VMError)
			if !ok {
				t.Fatalf("expected *VMError, got %T: %v", err, err)
			}
			if vme.Number != HResultServerXMLHTTPUnrecognizedScheme {
				t.Fatalf("expected HResult 0x80072EE6 (%d), got %d (0x%08X)",
					HResultServerXMLHTTPUnrecognizedScheme, vme.Number, uint32(vme.Number))
			}
			if !strings.Contains(vme.Description, "The URL does not use a recognized protocol") {
				t.Fatalf("unexpected description: %q", vme.Description)
			}
		})
	}

	// Test from ASP with On Error Resume Next
	aspSource := `<%
On Error Resume Next
Dim http
Set http = CreateObject("MSXML2.ServerXMLHTTP")
http.Open "GET", "/relative/url.asp", False
Response.Write Hex(Err.Number) & "|" & Err.Description
%>`
	output := runASPSource(t, aspSource, nil)
	expectedPrefix := "80072EE6|The URL does not use a recognized protocol"
	if !strings.HasPrefix(output, expectedPrefix) {
		t.Fatalf("expected %q, got %q", expectedPrefix, output)
	}

	// Test from ASP without On Error Resume Next (should fail execution)
	aspFailSource := `<%
Dim http
Set http = CreateObject("MSXML2.ServerXMLHTTP")
http.Open "GET", "/relative/url.asp", False
Response.Write "SHOULD_NOT_REACH"
%>`
	compiler := NewASPCompiler(aspFailSource)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}
	failVM := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	failHost := NewMockHost()
	var out bytes.Buffer
	failHost.SetOutput(&out)
	failVM.SetHost(failHost)
	err := failVM.Run()
	if err == nil {
		t.Fatalf("expected VM run error, but it succeeded: %s", out.String())
	}
	vme, ok := err.(*VMError)
	if !ok || vme.Number != HResultServerXMLHTTPUnrecognizedScheme {
		t.Fatalf("expected unhandled error 0x80072EE6, got: %v", err)
	}
}

// TestMSXMLServerXMLHTTPSendAfterInvalidOpen verifies that executing Send after an invalid Open raises 0x80004005.
func TestMSXMLServerXMLHTTPSendAfterInvalidOpen(t *testing.T) {
	vm := NewVM(nil, nil, 0)
	vm.host = &MockHost{}

	// Direct call test
	httpLib := NewMsXML2ServerXMLHTTP(vm)
	_, _ = httpLib.legacyCallMethod("open", "GET", "/invalid/relative")
	_, err := httpLib.legacyCallMethod("send")
	if err == nil {
		t.Fatal("expected send to fail after invalid open")
	}
	vme, ok := err.(*VMError)
	if !ok {
		t.Fatalf("expected *VMError, got %T: %v", err, err)
	}
	if vme.Number != HResultServerXMLHTTPUnspecifiedError {
		t.Fatalf("expected HResult 0x80004005 (%d), got %d (0x%08X)",
			HResultServerXMLHTTPUnspecifiedError, vme.Number, uint32(vme.Number))
	}

	// ASP test with On Error Resume Next
	aspSource := `<%
On Error Resume Next
Dim http
Set http = CreateObject("MSXML2.ServerXMLHTTP")
http.Open "GET", "/invalid/relative", False
Err.Clear
http.Send
Response.Write Hex(Err.Number) & "|" & Err.Description
%>`
	output := runASPSource(t, aspSource, nil)
	expectedPrefix := "80004005|Unspecified error"
	if !strings.HasPrefix(output, expectedPrefix) {
		t.Fatalf("expected %q, got %q", expectedPrefix, output)
	}
}

// TestMSXMLServerXMLHTTPSendNetworkFailures verifies that connection refused, DNS failure, and timeouts
// raise COM HRESULT 0x80072EFD and sanitize statusText.
func TestMSXMLServerXMLHTTPSendNetworkFailures(t *testing.T) {
	networkErrors := []struct {
		name string
		err  error
	}{
		{
			name: "connection_refused",
			err:  syscall.ECONNREFUSED,
		},
		{
			name: "dns_error",
			err:  &net.DNSError{Err: "no such host", Name: "nonexistent.example.invalid"},
		},
		{
			name: "timeout_error",
			err:  &net.OpError{Op: "dial", Net: "tcp", Err: errors.New("i/o timeout")},
		},
	}

	for _, tc := range networkErrors {
		t.Run(tc.name, func(t *testing.T) {
			vm := NewVM(nil, nil, 0)
			vm.host = &MockHost{}
			httpLib := NewMsXML2ServerXMLHTTP(vm)
			httpLib.SetTransport(&mockRoundTripper{
				roundTripFunc: func(req *http.Request) (*http.Response, error) {
					return nil, tc.err
				},
			})

			_, err := httpLib.legacyCallMethod("open", "GET", "http://example.com/test")
			if err != nil {
				t.Fatalf("open failed: %v", err)
			}

			_, sendErr := httpLib.legacyCallMethod("send")
			if sendErr == nil {
				t.Fatal("expected send to fail on network error")
			}
			vme, ok := sendErr.(*VMError)
			if !ok {
				t.Fatalf("expected *VMError, got %T: %v", sendErr, sendErr)
			}
			if vme.Number != HResultServerXMLHTTPCannotConnect {
				t.Fatalf("expected HResult 0x80072EFD (%d), got %d (0x%08X)",
					HResultServerXMLHTTPCannotConnect, vme.Number, uint32(vme.Number))
			}
			if !strings.Contains(vme.Description, "A connection with the server could not be established") {
				t.Fatalf("unexpected description: %q", vme.Description)
			}

			// Verify statusText does not leak Go error string
			if httpLib.statusText != "" {
				t.Fatalf("expected empty statusText on failure, got %q", httpLib.statusText)
			}
		})
	}

	// Real OS network failure test (closed port on 127.0.0.1)
	ln, err := net.Listen("tcp", "127.0.0.1:0")
	if err != nil {
		t.Fatalf("failed to listen on local port: %v", err)
	}
	closedAddr := ln.Addr().String()
	_ = ln.Close()

	aspSource := fmt.Sprintf(`<%s
On Error Resume Next
Dim http
Set http = Server.CreateObject("MSXML2.ServerXMLHTTP")
http.Open "GET", "http://%s/test", False
http.Send
Response.Write Hex(Err.Number) & "|" & Err.Description
%%>`, "", closedAddr)

	output := runASPSource(t, aspSource, nil)
	expectedPrefix := "80072EFD|A connection with the server could not be established"
	if !strings.HasPrefix(output, expectedPrefix) {
		t.Fatalf("expected %q, got %q", expectedPrefix, output)
	}
}

// TestMSXMLServerXMLHTTPStatusGetterAfterFailedSend verifies that reading .status after a failed send
// raises an exception and does not return 0.
func TestMSXMLServerXMLHTTPStatusGetterAfterFailedSend(t *testing.T) {
	vm := NewVM(nil, nil, 0)
	vm.host = &MockHost{}

	httpLib := NewMsXML2ServerXMLHTTP(vm)
	httpLib.SetTransport(&mockRoundTripper{
		roundTripFunc: func(req *http.Request) (*http.Response, error) {
			return nil, syscall.ECONNREFUSED
		},
	})

	_, _ = httpLib.legacyCallMethod("open", "GET", "http://example.com/test")
	_, _ = httpLib.legacyCallMethod("send")

	// Verify reading .status triggers an exception
	defer func() {
		r := recover()
		if r == nil {
			t.Fatal("expected reading status after failed send to raise an exception / panic")
		}
		vme, ok := r.(*VMError)
		if !ok {
			t.Fatalf("expected *VMError, got %T: %v", r, r)
		}
		if vme.Number != HResultServerXMLHTTPDataNotAvailable {
			t.Fatalf("expected HResult 0x8000000A (%d), got %d (0x%08X)",
				HResultServerXMLHTTPDataNotAvailable, vme.Number, uint32(vme.Number))
		}
	}()

	// This should panic / raise exception because ctx has no On Error Resume Next
	_ = httpLib.DispatchPropertyGet("status")
}

// TestMSXMLServerXMLHTTPStatusGetterInASP verifies ASP On Error Resume Next captures the .status getter exception.
func TestMSXMLServerXMLHTTPStatusGetterInASP(t *testing.T) {
	ln, err := net.Listen("tcp", "127.0.0.1:0")
	if err != nil {
		t.Fatalf("failed to listen on local port: %v", err)
	}
	closedAddr := ln.Addr().String()
	_ = ln.Close()

	aspSource := fmt.Sprintf(`<%s
On Error Resume Next
Dim http, st
Set http = Server.CreateObject("MSXML2.ServerXMLHTTP")
http.Open "GET", "http://%s/test", False
http.Send
Err.Clear
st = http.Status
Response.Write Hex(Err.Number) & "|" & Err.Description
%%>`, "", closedAddr)

	output := runASPSource(t, aspSource, nil)
	expectedPrefix := "8000000A|The data necessary to complete this operation is not yet available"
	if !strings.HasPrefix(output, expectedPrefix) {
		t.Fatalf("expected %q, got %q", expectedPrefix, output)
	}
}

// TestMSXMLServerXMLHTTPSuccess verifies happy path execution and properties.
func TestMSXMLServerXMLHTTPSuccess(t *testing.T) {
	ts := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		w.Header().Set("Content-Type", "text/plain")
		w.Header().Set("X-Custom-Header", "AxonTest")
		w.WriteHeader(http.StatusOK)
		_, _ = w.Write([]byte("response payload"))
	}))
	defer ts.Close()

	aspSource := fmt.Sprintf(`<%s
Dim http
Set http = Server.CreateObject("MSXML2.ServerXMLHTTP")
http.Open "GET", "%s", False
http.Send
Response.Write http.Status & "|" & http.StatusText & "|" & http.ResponseText & "|" & http.GetResponseHeader("X-Custom-Header")
%%>`, "", ts.URL)

	output := runASPSource(t, aspSource, nil)
	expected := "200|200 OK|response payload|AxonTest"
	if output != expected {
		t.Fatalf("expected %q, got %q", expected, output)
	}
}

