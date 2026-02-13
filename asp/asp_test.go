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
package asp

import (
	"os"
	"path/filepath"
	"strings"
	"testing"
)

// TestASPLexerBasic testa a análise léxica básica
func TestASPLexerBasic(t *testing.T) {
	code := `<% Dim x %>HTML`
	lexer := NewASPLexer(code)
	blocks := lexer.Tokenize()

	if len(blocks) != 2 {
		t.Errorf("Expected 2 blocks, got %d", len(blocks))
	}

	if blocks[0].Type != "asp" {
		t.Errorf("Expected first block to be 'asp', got '%s'", blocks[0].Type)
	}

	if blocks[1].Type != "html" {
		t.Errorf("Expected second block to be 'html', got '%s'", blocks[1].Type)
	}
}

// TestASPLexerMultipleBlocks testa múltiplos blocos
func TestASPLexerMultipleBlocks(t *testing.T) {
	code := `<% Dim x %>HTML1<% Dim y %>HTML2`
	lexer := NewASPLexer(code)
	blocks := lexer.Tokenize()

	if len(blocks) != 4 {
		t.Errorf("Expected 4 blocks, got %d", len(blocks))
	}
}

// TestASPLexerNoBlocks testa código sem blocos ASP
func TestASPLexerNoBlocks(t *testing.T) {
	code := `<html><body>Only HTML</body></html>`
	lexer := NewASPLexer(code)
	blocks := lexer.Tokenize()

	if len(blocks) != 1 || blocks[0].Type != "html" {
		t.Errorf("Expected one HTML block")
	}
}

// TestASPLexerUnclosedBlock testa bloco não fechado
func TestASPLexerUnclosedBlock(t *testing.T) {
	code := `<% Dim x`
	lexer := NewASPLexer(code)
	blocks := lexer.Tokenize()

	// Bloco não fechado deve ser ignorado
	if len(blocks) > 0 && blocks[0].Type == "asp" {
		t.Errorf("Unclosed block should not be treated as ASP")
	}
}

// TestASPParserSimple testa parsing simples
func TestASPParserSimple(t *testing.T) {
	code := `<% Dim message %><html><body><%= message %></body></html>`
	parser := NewASPParser(code)
	result, err := parser.Parse()

	if err != nil {
		t.Errorf("Parse error: %v", err)
	}

	if len(result.Blocks) == 0 {
		t.Errorf("Expected blocks, got none")
	}
}

// TestExtractHTMLOnly testa extração de HTML
func TestExtractHTMLOnly(t *testing.T) {
	code := `<% Dim x %><html>Content</html><% x = 5 %>`
	html := ExtractHTMLOnly(code)

	if !strings.Contains(html, "<html>") || !strings.Contains(html, "</html>") {
		t.Errorf("HTML not properly extracted")
	}

	if strings.Contains(html, "Dim x") {
		t.Errorf("VB code should not be in extracted HTML")
	}
}

// TestExtractVBScriptOnly testa extração de VBScript
func TestExtractVBScriptOnly(t *testing.T) {
	code := `<% Dim x %><html>Content</html><% x = 5 %>`
	vb := ExtractVBScriptOnly(code)

	if !strings.Contains(vb, "Dim x") {
		t.Errorf("VB code 'Dim x' not found in extraction")
	}

	if strings.Contains(vb, "<html>") {
		t.Errorf("HTML should not be in extracted VB code")
	}
}

// TestServerObject testa o objeto Server
func TestServerObject(t *testing.T) {
	server := NewServerObject()

	if server.GetName() != "Server" {
		t.Errorf("Server name should be 'Server'")
	}

	// Testa URLEncode
	encoded, _ := server.CallMethod("URLEncode", "hello world")
	if encoded == "" {
		t.Errorf("URLEncode should return a value")
	}
}

// TestRequestObject testa o objeto Request
func TestRequestObject(t *testing.T) {
	request := NewRequestObject()

	if request.GetName() != "Request" {
		t.Errorf("Request name should be 'Request'")
	}

	request.SetProperty("test_prop", "test_value")

	if request.GetProperty("test_prop") != "test_value" {
		t.Errorf("Property not set correctly")
	}
}

// TestResponseObject testa o objeto Response
func TestResponseObject(t *testing.T) {
	response := NewResponseObject()

	if response.GetName() != "Response" {
		t.Errorf("Response name should be 'Response'")
	}

	response.CallMethod("Write", "Hello")
	buffer := response.GetBuffer()

	if buffer != "Hello" {
		t.Errorf("Response buffer should contain 'Hello', got '%s'", buffer)
	}
}

// TestASPContext testa o contexto ASP
func TestASPContext(t *testing.T) {
	context := NewASPContext()

	if context.Server == nil {
		t.Errorf("Server object should be initialized")
	}

	if context.Request == nil {
		t.Errorf("Request object should be initialized")
	}

	if context.Response == nil {
		t.Errorf("Response object should be initialized")
	}

	if context.Session == nil {
		t.Errorf("Session object should be initialized")
	}

	if context.Application == nil {
		t.Errorf("Application object should be initialized")
	}
}

// TestASPValidator testa validação
func TestASPValidator(t *testing.T) {
	validator := NewASPValidator()

	validCode := `<% Dim x %>`
	valid, msgs := validator.Validate(validCode)

	if !valid {
		t.Errorf("Simple VB code should be valid. Messages: %v", msgs)
	}
}

func TestASPParser_CKEditorInclude(t *testing.T) {
	wd, err := os.Getwd()
	if err != nil {
		t.Fatalf("failed to get working directory: %v", err)
	}
	ckeditorPath := filepath.Join(wd, "..", "www", "QuickerSite-test", "asp", "includes", "ckeditor.asp")
	content, err := ReadFileText(ckeditorPath)
	if err != nil {
		t.Fatalf("failed to read ckeditor.asp: %v", err)
	}

	parser := NewASPParser(content)
	result, err := parser.Parse()
	if err != nil {
		t.Fatalf("parse error: %v", err)
	}
	if len(result.Errors) > 0 {
		t.Fatalf("asp parse errors: %v", result.Errors[0])
	}
}

func TestASPParser_QuickerSiteDefault(t *testing.T) {
	wd, err := os.Getwd()
	if err != nil {
		t.Fatalf("failed to get working directory: %v", err)
	}
	rootDir := filepath.Join(wd, "..", "www")
	defaultPath := filepath.Join(rootDir, "QuickerSite-test", "default.asp")
	content, err := ReadFileText(defaultPath)
	if err != nil {
		t.Fatalf("failed to read default.asp: %v", err)
	}

	options := NewASPParsingOptions()
	resolved, result, err := ParseWithCache(content, defaultPath, rootDir, options)
	if err != nil {
		t.Fatalf("parse error: %v", err)
	}
	if resolved == "" {
		t.Fatalf("resolved content is empty")
	}
	if len(result.Errors) > 0 {
		t.Fatalf("asp parse errors: %v", result.Errors[0])
	}
}

// TestASPFormatter testa formatação
func TestASPFormatter(t *testing.T) {
	formatter := NewASPFormatter(2)
	code := `<% Dim x %>`
	formatted := formatter.Format(code)

	if !strings.Contains(formatted, "<%") || !strings.Contains(formatted, "%>") {
		t.Errorf("Formatted code should contain ASP delimiters")
	}
}

// TestASPError testa manipulação de erros
func TestASPError(t *testing.T) {
	err := NewASPError("TEST001", "Test error message", 10, 5)

	if err.Code != "TEST001" {
		t.Errorf("Error code mismatch")
	}

	if err.Line != 10 {
		t.Errorf("Error line mismatch")
	}
}

// TestASPErrorHandler testa manipulador de erros
func TestASPErrorHandler(t *testing.T) {
	handler := NewASPErrorHandler()

	if handler.HasErrors() {
		t.Errorf("Handler should start with no errors")
	}

	err := NewASPError("TEST", "Test", 1, 1)
	handler.AddError(err)

	if !handler.HasErrors() {
		t.Errorf("Handler should have errors after adding one")
	}

	if handler.GetErrorCount() != 1 {
		t.Errorf("Error count should be 1")
	}
}

// BenchmarkASPLexer benchmark para análise léxica
func BenchmarkASPLexer(b *testing.B) {
	code := `
<% 
	Dim name
	name = "John"
%>
<html>
<body>
	<h1><%= name %></h1>
</body>
</html>
<% 
	Response.Write(name)
%>
`

	b.ResetTimer()
	for i := 0; i < b.N; i++ {
		lexer := NewASPLexer(code)
		lexer.Tokenize()
	}
}

// BenchmarkASPParser benchmark para parsing
func BenchmarkASPParser(b *testing.B) {
	code := `
<% 
	Dim name
	name = "John"
%>
<html>
<body>
	<h1><%= name %></h1>
</body>
</html>
<% 
	Response.Write(name)
%>
`

	b.ResetTimer()
	for i := 0; i < b.N; i++ {
		parser := NewASPParser(code)
		parser.Parse()
	}
}

// TestASPDirective tests ASP directive parsing
func TestASPDirective(t *testing.T) {
	code := `<%@ Language=VBScript %>
<% Dim x %>
<html>Content</html>`

	lexer := NewASPLexer(code)
	blocks := lexer.Tokenize()

	if len(blocks) < 1 {
		t.Fatalf("Expected at least 1 block, got %d", len(blocks))
	}

	// First block should be directive
	if blocks[0].Type != "directive" {
		t.Errorf("Expected first block to be 'directive', got '%s'", blocks[0].Type)
	}

	// Check directive attributes
	if blocks[0].Attributes == nil {
		t.Errorf("Expected directive attributes, got nil")
	} else if lang, exists := blocks[0].Attributes["Language"]; !exists || lang != "VBScript" {
		t.Errorf("Expected Language=VBScript, got %v", blocks[0].Attributes)
	}
}

// TestASPDirectiveWithQuotes tests directive with quoted values
func TestASPDirectiveWithQuotes(t *testing.T) {
	code := `<%@ Language="VBScript" CodePage="65001" %>`

	lexer := NewASPLexer(code)
	blocks := lexer.Tokenize()

	if len(blocks) < 1 {
		t.Fatalf("Expected at least 1 block, got %d", len(blocks))
	}

	// Check directive attributes
	if blocks[0].Attributes == nil {
		t.Fatalf("Expected directive attributes, got nil")
	}

	if lang, exists := blocks[0].Attributes["Language"]; !exists || lang != "VBScript" {
		t.Errorf("Expected Language=VBScript, got %v", blocks[0].Attributes)
	}

	if cp, exists := blocks[0].Attributes["CodePage"]; !exists || cp != "65001" {
		t.Errorf("Expected CodePage=65001, got %v", blocks[0].Attributes)
	}
}

// TestASPParserWithDirective tests that parser handles directives correctly
func TestASPParserWithDirective(t *testing.T) {
	code := `<%@ Language=VBScript %>
<% Dim message
   message = "Hello" %>
<html><body><%= message %></body></html>`

	parser := NewASPParser(code)
	result, err := parser.Parse()

	if err != nil {
		t.Errorf("Parse error: %v", err)
	}

	if len(result.Blocks) == 0 {
		t.Errorf("Expected blocks, got none")
	}

	// Verify directive block exists
	hasDirective := false
	for _, block := range result.Blocks {
		if block.Type == "directive" {
			hasDirective = true
			break
		}
	}

	if !hasDirective {
		t.Errorf("Expected to find directive block")
	}
}

func TestASPParserSuppressesStructuralWhitespaceBetweenASPBlocks(t *testing.T) {
	code := "<% Dim firstValue %>\r\n\r\n\r\n<% Dim secondValue %>"

	parser := NewASPParser(code)
	result, err := parser.Parse()
	if err != nil {
		t.Fatalf("Parse error: %v", err)
	}

	if len(result.Errors) > 0 {
		t.Fatalf("Unexpected parse errors: %v", result.Errors[0])
	}

	if strings.Contains(result.CombinedVBCode, "Response.Write") {
		t.Fatalf("Expected no Response.Write for structural whitespace, got: %s", result.CombinedVBCode)
	}
}

func TestASPParserKeepsWhitespaceInsideRealHTMLBlocks(t *testing.T) {
	code := "<div> \n</div>"

	parser := NewASPParser(code)
	result, err := parser.Parse()
	if err != nil {
		t.Fatalf("Parse error: %v", err)
	}

	if len(result.Errors) > 0 {
		t.Fatalf("Unexpected parse errors: %v", result.Errors[0])
	}

	if !strings.Contains(result.CombinedVBCode, "Response.Write") {
		t.Fatalf("Expected Response.Write for real HTML output")
	}
}

func TestNormalizeBoundaryHTMLContent_TrimsLeadingCRLFBeforeHTML(t *testing.T) {
	blocks := []*CodeBlock{
		{Type: "asp", Content: "Dim firstValue"},
		{Type: "html", Content: "\r\n\r\n<root/>"},
	}

	normalized := normalizeBoundaryHTMLContent(blocks, 1, blocks[1].Content)
	if normalized != "<root/>" {
		t.Fatalf("expected leading boundary CRLF to be trimmed, got %q", normalized)
	}
}

func TestNormalizeBoundaryHTMLContent_TrimsTrailingCRLFBeforeASP(t *testing.T) {
	blocks := []*CodeBlock{
		{Type: "html", Content: "<root/>\r\n\r\n"},
		{Type: "asp", Content: "Dim firstValue"},
	}

	normalized := normalizeBoundaryHTMLContent(blocks, 0, blocks[0].Content)
	if normalized != "<root/>" {
		t.Fatalf("expected trailing boundary CRLF to be trimmed, got %q", normalized)
	}
}
