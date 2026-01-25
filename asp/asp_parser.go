/*
 * AxonASP Server - Version 1.0
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
	"fmt"
	vb "go-asp/vbscript"
	"go-asp/vbscript/ast"
	"strings"
)

// ASPParserResult contém o resultado da análise de código ASP
type ASPParserResult struct {
	Blocks          []*CodeBlock         // All ASP/HTML blocks in order
	VBPrograms      map[int]*ast.Program // Parsed VBScript programs keyed by block index
	CombinedProgram *ast.Program         // Full VBScript program with HTML injected as Response.Write
	CombinedVBCode  string               // Combined VBScript source code
	Errors          []error              // Parse errors
	HTMLContent     []string             // HTML content in order
}

// ASPParser realiza análise sintática de código ASP
// Usa o Parser do VBScript-Go para processar blocos de código ASP
type ASPParser struct {
	lexer     *ASPLexer
	options   *ASPParsingOptions
	vbOptions *vb.ParsingOptions
}

// ASPParsingOptions contém opções para análise de ASP
type ASPParsingOptions struct {
	SaveComments      bool
	StrictMode        bool
	AllowImplicitVars bool
	DebugMode         bool
}

// NewASPParsingOptions cria opções padrão
func NewASPParsingOptions() *ASPParsingOptions {
	return &ASPParsingOptions{
		SaveComments:      false,
		StrictMode:        false,
		AllowImplicitVars: true,
		DebugMode:         false,
	}
}

// NewASPParser cria um novo parser ASP
func NewASPParser(code string) *ASPParser {
	return NewASPParserWithOptions(code, NewASPParsingOptions())
}

// NewASPParserWithOptions cria um novo parser ASP com opções customizadas
func NewASPParserWithOptions(code string, options *ASPParsingOptions) *ASPParser {
	if options == nil {
		options = NewASPParsingOptions()
	}

	vbOptions := &vb.ParsingOptions{
		SaveComments: options.SaveComments,
	}

	return &ASPParser{
		lexer:     NewASPLexer(code),
		options:   options,
		vbOptions: vbOptions,
	}
}

// Parse realiza análise completa do código ASP
func (ap *ASPParser) Parse() (*ASPParserResult, error) {
	result := &ASPParserResult{
		Blocks:      make([]*CodeBlock, 0),
		VBPrograms:  make(map[int]*ast.Program),
		Errors:      make([]error, 0),
		HTMLContent: make([]string, 0),
	}

	// Tokeniza o código ASP
	blocks := ap.lexer.Tokenize()
	result.Blocks = blocks

	// Processa cada bloco
	for i, block := range blocks {
		switch block.Type {
		case "directive":
			// ASP directives like <%@ Language=VBScript %>
			// These are processed but don't generate code
			// They are used for configuration (Language, CodePage, etc.)
			// No action needed here - directives are parsed in the lexer
		case "asp":
			// Attempt to parse individual VBScript block without failing the whole page
			program, err := ap.parseVBBlock(block.Content)
			if err != nil {
				if ap.options.DebugMode {
					fmt.Printf("[ASP Parser Error] Line %d: %v\n", block.Line, err)
				}
				break
			}
			result.VBPrograms[i] = program
		case "html":
			// Armazena conteúdo HTML
			result.HTMLContent = append(result.HTMLContent, block.Content)
		}
	}

	// Build a single VBScript program that includes HTML as Response.Write calls
	combinedVB := buildCombinedVBScript(blocks)
	result.CombinedVBCode = combinedVB

	if strings.TrimSpace(combinedVB) != "" {
		program, err := ap.parseVBBlock(combinedVB)
		if err != nil {
			parseErr := fmt.Errorf("Parse error in combined ASP script: %v", err)
			result.Errors = append(result.Errors, parseErr)
			if ap.options.DebugMode {
				fmt.Printf("[ASP Parser Error] Combined: %v\n", err)
			}
		} else {
			result.CombinedProgram = program
			// Store under a reserved index so existing consumers can iterate
			result.VBPrograms[-1] = program
			// Combined program parsed successfully; suppress any block-level parse noise
			result.Errors = nil
		}
	}

	return result, nil
}

// parseVBBlock realiza parse de um bloco de código VBScript
func (ap *ASPParser) parseVBBlock(code string) (program *ast.Program, err error) {
	// Remove comentários vazios e espaços em branco
	trimmedCode := strings.TrimSpace(code)

	if trimmedCode == "" {
		// Bloco vazio é válido
		return &ast.Program{
				OptionExplicit: false,
				OptionCompare:  ast.OptionCompareBinary,
				Body:           []ast.Statement{},
			},
			nil
	}

	// Pre-process colons to handle multi-statement lines
	// The VBScript parser might panic on colons, so we convert them to newlines
	// processedCode := preProcessColons(trimmedCode)
	processedCode := trimmedCode

	// Usa o parser do VBScript-Go
	parser := vb.NewParserWithOptions(processedCode, ap.vbOptions)

	// Faz o parse e captura possíveis panics
	defer func() {
		if r := recover(); r != nil {
			if e, ok := r.(error); ok {
				err = e
			} else {
				err = fmt.Errorf("panic during parse: %v", r)
			}
		}
	}()

	program = parser.Parse()
	return program, nil
}

// preProcessColons replaces colons with newlines outside of strings and comments
func preProcessColons(code string) string {
	var sb strings.Builder
	inString := false
	inComment := false
	inDateLiteral := false

	for i := 0; i < len(code); i++ {
		char := code[i]

		if char == '\n' || char == '\r' {
			inComment = false
			inString = false      // Strings don't span lines in VBScript
			inDateLiteral = false // Date literals don't span lines
			sb.WriteByte(char)
			continue
		}

		if inComment {
			sb.WriteByte(char)
			continue
		}

		if char == '"' {
			if !inDateLiteral {
				inString = !inString
			}
			sb.WriteByte(char)
			continue
		}

		if inString {
			sb.WriteByte(char)
			continue
		}

		// Check for date literal delimiter
		if char == '#' {
			inDateLiteral = !inDateLiteral
			sb.WriteByte(char)
			continue
		}

		if inDateLiteral {
			sb.WriteByte(char)
			continue
		}

		// Check for comment start
		if char == '\'' {
			inComment = true
			sb.WriteByte(char)
			continue
		}
		// Check for REM comment
		if (char == 'R' || char == 'r') && i+3 < len(code) {
			if strings.EqualFold(code[i:i+4], "REM ") {
				inComment = true
				sb.WriteString(code[i : i+4])
				i += 3
				continue
			}
		}

		// Replace colon with newline
		if char == ':' {
			sb.WriteByte('\n')
		} else {
			sb.WriteByte(char)
		}
	}

	return sb.String()
}

// buildCombinedVBScript merges ASP blocks and HTML into a single VBScript program
// HTML segments are converted into Response.Write calls so they participate in control flow
func buildCombinedVBScript(blocks []*CodeBlock) string {
	var sb strings.Builder

	optionDirectives := make([]string, 0)
	adjustedContents := make(map[int]string, len(blocks))

	for idx, block := range blocks {
		if block.Type != "asp" {
			continue
		}

		content := block.Content
		if strings.TrimSpace(content) == "" {
			adjustedContents[idx] = ""
			continue
		}

		lines := strings.Split(content, "\n")
		remainingLines := make([]string, 0, len(lines))
		seenExecutable := false

		for _, line := range lines {
			trimmed := strings.TrimSpace(line)
			lower := strings.ToLower(trimmed)
			isComment := strings.HasPrefix(lower, "rem ") || strings.HasPrefix(trimmed, "'")

			if !seenExecutable {
				if trimmed == "" || isComment {
					remainingLines = append(remainingLines, line)
					continue
				}

				if strings.HasPrefix(lower, "option explicit") || strings.HasPrefix(lower, "option base") {
					optionDirectives = append(optionDirectives, trimmed)
					continue
				}

				seenExecutable = true
			}

			remainingLines = append(remainingLines, line)
		}

		adjustedContents[idx] = strings.Join(remainingLines, "\n")
	}

	for _, directive := range optionDirectives {
		sb.WriteString(directive)
		sb.WriteString("\n")
	}

	for idx, block := range blocks {
		switch block.Type {
		case "asp":
			content := block.Content
			if adjusted, ok := adjustedContents[idx]; ok {
				content = adjusted
			}
			content = strings.TrimSpace(content)
			if content == "" {
				continue
			}
			sb.WriteString(content)
			sb.WriteString("\n")
		case "html":
			htmlWrite := htmlToVBWrite(block.Content)
			if htmlWrite == "" {
				continue
			}
			sb.WriteString(htmlWrite)
			sb.WriteString("\n")
		}
	}

	return sb.String()
}

// htmlToVBWrite converts raw HTML into a VBScript Response.Write statement
// New lines are preserved using vbCrLf and quotes are doubled for VBScript string literals
func htmlToVBWrite(html string) string {
	if html == "" {
		return ""
	}

	normalized := strings.ReplaceAll(html, "\r\n", "\n")
	lines := strings.Split(normalized, "\n")
	parts := make([]string, 0, len(lines)*2)

	for i, line := range lines {
		escaped := strings.ReplaceAll(line, "\"", "\"\"")
		parts = append(parts, fmt.Sprintf("\"%s\"", escaped))
		if i < len(lines)-1 {
			parts = append(parts, "vbCrLf")
		}
	}

	return "Response.Write " + strings.Join(parts, " & ")
}

// GetVBProgramsFromResult retorna os programas VBScript de um resultado
func (ap *ASPParser) GetVBProgramsFromResult(result *ASPParserResult) map[int]*ast.Program {
	if result != nil {
		return result.VBPrograms
	}
	return make(map[int]*ast.Program)
}

// ExtractVBScriptCode extrai apenas o código VBScript dos blocos ASP
func (ap *ASPParser) ExtractVBScriptCode(separator string) string {
	blocks := ap.lexer.Tokenize()
	var vbCode []string

	for _, block := range blocks {
		if block.Type == "asp" {
			vbCode = append(vbCode, block.Content)
		}
	}

	if separator == "" {
		separator = "\n"
	}

	return strings.Join(vbCode, separator)
}

// Reset reinicia o parser
func (ap *ASPParser) Reset() {
	ap.lexer.Reset()
}
