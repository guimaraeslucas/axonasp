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
	"fmt"
	"os"
	"strings"

	"g3pix.com.br/axonasp/experimental"
)

// ASPExecutor executa código ASP parseado em um contexto
type ASPExecutor struct {
	context *ASPContext
	parser  *ASPParser
}

// NewASPExecutor cria um novo executor ASP
func NewASPExecutor() *ASPExecutor {
	return &ASPExecutor{
		context: NewASPContext(),
	}
}

// Execute executa código ASP e retorna o resultado
func (ae *ASPExecutor) Execute(aspCode string) (string, error) {
	// Check for VM enablement via environment variable
	if os.Getenv("AXONASP_VM") == "1" {
		return ae.ExecuteVM(aspCode)
	}

	parser := NewASPParser(aspCode)
	result, err := parser.Parse()

	if err != nil {
		return "", err
	}

	if len(result.Errors) > 0 {
		return "", fmt.Errorf("erros durante análise: %v", result.Errors[0])
	}

	ae.parser = parser

	// Executa todos os blocos VBScript em ordem
	for i, block := range result.Blocks {
		if block.Type == "asp" {
			if program, exists := result.VBPrograms[i]; exists && program != nil {
				// Aqui você poderia executar o programa VBScript
				// Por enquanto, apenas coletamos a saída da Response
				_ = program
			}
		}
	}

	return ae.context.Response.GetBuffer(), nil
}

// ExecuteVM executes ASP code using the experimental Bytecode VM
func (ae *ASPExecutor) ExecuteVM(aspCode string) (string, error) {
	parser := NewASPParser(aspCode)
	result, err := parser.Parse()
	if err != nil {
		return "", err
	}

	compiler := experimental.NewCompiler()

	// Compile all ASP blocks
	for i, block := range result.Blocks {
		switch block.Type {
case "asp":
			if program, exists := result.VBPrograms[i]; exists && program != nil {
				if err := compiler.Compile(program); err != nil {
					return "", fmt.Errorf("VM compilation error: %v", err)
				}
			}
		case "html":
			// Handle HTML blocks? VM doesn't handle HTML directly usually.
			// Standard ASP emits HTML via Response.Write.
			// For now, we ignore HTML blocks in VM execution or manual write?
			// Let's manually append to buffer for simplicity in this experimental phase.
			ae.context.Response.buffer += block.Content
		}
	}

	vm := experimental.NewVM(compiler.MainFunction(), &ASPHostAdapter{Ctx: ae.context})
	if err := vm.Run(); err != nil {
		return "", fmt.Errorf("VM runtime error: %v", err)
	}

	return ae.context.Response.GetBuffer(), nil
}

// ASPHostAdapter adapts ASPContext to experimental.HostEnvironment
type ASPHostAdapter struct {
	Ctx *ASPContext
}

func (a *ASPHostAdapter) GetVariable(name string) (interface{}, bool) {
	// Check variables
	name = strings.ToLower(name)
	if val, ok := a.Ctx.Variables[name]; ok {
		return val, true
	}
	// Check objects?
	return nil, false
}

func (a *ASPHostAdapter) SetVariable(name string, value interface{}) error {
	name = strings.ToLower(name)
	a.Ctx.Variables[name] = value
	return nil
}

func (a *ASPHostAdapter) CallFunction(name string, args []interface{}) (interface{}, error) {
	// Basic implementation for testing
	name = strings.ToLower(name)
	switch name {
	case "response.write", "responsewrite":
		if len(args) > 0 {
			a.Ctx.Response.buffer += fmt.Sprintf("%v", args[0])
		}
		return nil, nil
	case "document.write", "documentwrite":
		if len(args) > 0 {
			a.Ctx.Response.buffer += fmt.Sprintf("%v", args[0])
		}
		return nil, nil
	case "server.createobject", "servercreateobject":
		return a.Ctx.Server.CallMethod("CreateObject", args...)
	}
	return nil, fmt.Errorf("function not implemented in ASPHostAdapter: %s", name)
}

func (a *ASPHostAdapter) CreateObject(progID string) (interface{}, error) {
	return nil, fmt.Errorf("CreateObject not implemented in ASPHostAdapter")
}

// GetContext retorna o contexto de execução atual
func (ae *ASPExecutor) GetContext() *ASPContext {
	return ae.context
}

// ExecuteFile executa um arquivo ASP (simulado)
func (ae *ASPExecutor) ExecuteFile(filePath string, fileContent string) (string, error) {
	return ae.Execute(fileContent)
}

// BuildASPDocument cria um documento ASP completo com HTML e VB
func BuildASPDocument(htmlParts []string, vbParts []string) string {
	if len(htmlParts) == 0 && len(vbParts) == 0 {
		return ""
	}

	result := strings.Builder{}

	for i := 0; i < len(htmlParts); i++ {
		if htmlParts[i] != "" {
			result.WriteString(htmlParts[i])
		}
		if i < len(vbParts) && vbParts[i] != "" {
			result.WriteString("<%\n")
			result.WriteString(vbParts[i])
			result.WriteString("\n%>")
		}
	}

	return result.String()
}

// ASPValidator valida ASP sem executar
type ASPValidator struct {
	parser *ASPParser
}

// NewASPValidator cria um novo validador ASP
func NewASPValidator() *ASPValidator {
	return &ASPValidator{}
}

// Validate valida código ASP
func (av *ASPValidator) Validate(aspCode string) (bool, []string) {
	parser := NewASPParser(aspCode)
	result, err := parser.Parse()

	var messages []string

	if err != nil {
		messages = append(messages, fmt.Sprintf("Erro crítico: %v", err))
		return false, messages
	}

	if len(result.Errors) > 0 {
		for _, e := range result.Errors {
			messages = append(messages, fmt.Sprintf("Erro de parse: %v", e))
		}
	}

	if len(messages) > 0 {
		return false, messages
	}

	return true, []string{"Código ASP válido"}
}

// ASPFormatter formata código ASP
type ASPFormatter struct {
	indentSize int
}

// NewASPFormatter cria um novo formatador
func NewASPFormatter(indentSize int) *ASPFormatter {
	if indentSize <= 0 {
		indentSize = 2
	}
	return &ASPFormatter{
		indentSize: indentSize,
	}
}

// Format formata código ASP
func (af *ASPFormatter) Format(aspCode string) string {
	lexer := NewASPLexer(aspCode)
	blocks := lexer.Tokenize()

	result := strings.Builder{}

	for _, block := range blocks {
		switch block.Type {
case "html":
			result.WriteString(block.Content)
		case "asp":
			result.WriteString("<%\n")
			// Formata o conteúdo VBScript
			formattedVB := af.formatVBContent(block.Content)
			result.WriteString(formattedVB)
			result.WriteString("\n%>")
		}
	}

	return result.String()
}

// formatVBContent formata o conteúdo VBScript
func (af *ASPFormatter) formatVBContent(vbCode string) string {
	lines := strings.Split(vbCode, "\n")
	var result []string

	for _, line := range lines {
		trimmed := strings.TrimSpace(line)
		if trimmed != "" {
			result = append(result, "\t"+trimmed)
		}
	}

	return strings.Join(result, "\n")
}
