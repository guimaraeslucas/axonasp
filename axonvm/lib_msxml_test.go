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
