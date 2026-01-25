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
	"strings"
)

// ASPVBIntegration fornece integração com VBScript-Go
// Mantém compatibilidade total com o parser VBScript original
type ASPVBIntegration struct {
	aspParser *ASPParser
	context   *ASPContext
}

// NewASPVBIntegration cria uma nova integração ASP-VB
func NewASPVBIntegration() *ASPVBIntegration {
	return &ASPVBIntegration{
		context: NewASPContext(),
	}
}

// ParseASPFile analisa um arquivo ASP mantendo compatibilidade com VBScript-Go
func (av *ASPVBIntegration) ParseASPFile(fileName string, content string) (*ASPParserResult, error) {
	av.aspParser = NewASPParser(content)
	result, err := av.aspParser.Parse()
	return result, err
}

// GetVBScriptCodeForVBScriptGo extrai código VBScript compatível com VBScript-Go
func (av *ASPVBIntegration) GetVBScriptCodeForVBScriptGo() string {
	// Combina todos os blocos VBScript em um único código
	// que pode ser analisado pelo VBScript-Go parser diretamente
	if av.aspParser == nil {
		return ""
	}

	return av.aspParser.ExtractVBScriptCode("\n\n")
}

// GetExecutionContext retorna o contexto de execução do ASP
func (av *ASPVBIntegration) GetExecutionContext() *ASPContext {
	return av.context
}

// ExecuteWithContext executa código ASP com um contexto específico
func (av *ASPVBIntegration) ExecuteWithContext(aspCode string, customContext *ASPContext) (string, error) {
	if customContext != nil {
		av.context = customContext
	}

	executor := NewASPExecutor()
	executor.context = av.context

	output, err := executor.Execute(aspCode)
	return output, err
}

// GetCompatibilityInfo fornece informações sobre compatibilidade com VBScript-Go
func (av *ASPVBIntegration) GetCompatibilityInfo() map[string]interface{} {
	return map[string]interface{}{
		"asp_module_version":  "1.0.0",
		"vbscript_compatible": true,
		"asp_objects_supported": []string{
			"Server",
			"Request",
			"Response",
			"Session",
			"Application",
		},
		"supported_delimiters": []string{
			"<% %>",
		},
		"features": []string{
			"ASP code extraction",
			"VBScript parsing",
			"Error handling",
			"Code validation",
			"Formatting",
			"ASP object stubs",
		},
	}
}

// ASPCodeAnalyzer traz análise profunda de código ASP
type ASPCodeAnalyzer struct {
	result *ASPParserResult
}

// NewASPCodeAnalyzer cria um novo analisador
func NewASPCodeAnalyzer() *ASPCodeAnalyzer {
	return &ASPCodeAnalyzer{}
}

// Analyze analisa código ASP e retorna informações detalhadas
func (aca *ASPCodeAnalyzer) Analyze(aspCode string) map[string]interface{} {
	parser := NewASPParser(aspCode)
	result, _ := parser.Parse()
	aca.result = result

	analysis := make(map[string]interface{})

	// Contagem de blocos
	aspBlockCount := 0
	htmlBlockCount := 0
	totalContent := 0

	for _, block := range result.Blocks {
		if block.Type == "asp" {
			aspBlockCount++
			totalContent += len(block.Content)
		} else if block.Type == "html" {
			htmlBlockCount++
			totalContent += len(block.Content)
		}
	}

	analysis["total_blocks"] = len(result.Blocks)
	analysis["asp_blocks"] = aspBlockCount
	analysis["html_blocks"] = htmlBlockCount
	analysis["total_content_length"] = totalContent

	// Detecção de padrões
	patterns := aca.detectPatterns(aspCode)
	analysis["patterns_detected"] = patterns

	// Erro count
	analysis["parse_errors"] = len(result.Errors)

	// Complexidade
	analysis["complexity"] = aca.calculateComplexity(result)

	return analysis
}

// detectPatterns detecta padrões comuns em código ASP
func (aca *ASPCodeAnalyzer) detectPatterns(aspCode string) []string {
	var patterns []string

	if strings.Contains(aspCode, "Response.Write") {
		patterns = append(patterns, "Response.Write")
	}

	if strings.Contains(aspCode, "Request.Form") {
		patterns = append(patterns, "Form handling")
	}

	if strings.Contains(aspCode, "Request.QueryString") {
		patterns = append(patterns, "Query string handling")
	}

	if strings.Contains(aspCode, "Session(") {
		patterns = append(patterns, "Session variables")
	}

	if strings.Contains(aspCode, "Application(") {
		patterns = append(patterns, "Application variables")
	}

	if strings.Contains(aspCode, "CreateObject") {
		patterns = append(patterns, "COM object creation")
	}

	if strings.Contains(aspCode, "Set conn") || strings.Contains(aspCode, "ADODB") {
		patterns = append(patterns, "Database connection")
	}

	if strings.Contains(aspCode, "If") && strings.Contains(aspCode, "End If") {
		patterns = append(patterns, "Conditional logic")
	}

	if strings.Contains(aspCode, "Do") && strings.Contains(aspCode, "Loop") {
		patterns = append(patterns, "Do loop")
	}

	if strings.Contains(aspCode, "For") && strings.Contains(aspCode, "Next") {
		patterns = append(patterns, "For loop")
	}

	return patterns
}

// calculateComplexity calcula complexidade do código
func (aca *ASPCodeAnalyzer) calculateComplexity(result *ASPParserResult) string {
	blockCount := len(result.Blocks)

	switch {
	case blockCount <= 3:
		return "Low"
	case blockCount <= 10:
		return "Medium"
	case blockCount <= 20:
		return "High"
	default:
		return "Very High"
	}
}

// PrintAnalysis imprime análise em formato legível
func (aca *ASPCodeAnalyzer) PrintAnalysis(aspCode string) {
	analysis := aca.Analyze(aspCode)

	fmt.Println("=== ASP Code Analysis ===")
	fmt.Printf("Total Blocks: %d\n", analysis["total_blocks"])
	fmt.Printf("ASP Blocks: %d\n", analysis["asp_blocks"])
	fmt.Printf("HTML Blocks: %d\n", analysis["html_blocks"])
	fmt.Printf("Total Content Length: %d\n", analysis["total_content_length"])
	fmt.Printf("Parse Errors: %d\n", analysis["parse_errors"])
	fmt.Printf("Complexity: %s\n", analysis["complexity"])

	if patterns, ok := analysis["patterns_detected"].([]string); ok && len(patterns) > 0 {
		fmt.Println("Patterns Detected:")
		for _, pattern := range patterns {
			fmt.Printf("  - %s\n", pattern)
		}
	}
}

// ASPStatistics coleta estatísticas de código ASP
type ASPStatistics struct {
	TotalFiles      int
	TotalBlocks     int
	TotalASPBlocks  int
	TotalHTMLBlocks int
	TotalErrors     int
	AverageSize     float64
}

// NewASPStatistics cria novo coletor de estatísticas
func NewASPStatistics() *ASPStatistics {
	return &ASPStatistics{}
}

// AnalyzeFile analisa um arquivo e atualiza estatísticas
func (as *ASPStatistics) AnalyzeFile(aspCode string) {
	parser := NewASPParser(aspCode)
	result, _ := parser.Parse()

	as.TotalFiles++
	as.TotalBlocks += len(result.Blocks)
	as.TotalErrors += len(result.Errors)

	for _, block := range result.Blocks {
		if block.Type == "asp" {
			as.TotalASPBlocks++
		} else if block.Type == "html" {
			as.TotalHTMLBlocks++
		}
	}
}

// GetStatistics retorna as estatísticas coletadas
func (as *ASPStatistics) GetStatistics() map[string]interface{} {
	return map[string]interface{}{
		"total_files":       as.TotalFiles,
		"total_blocks":      as.TotalBlocks,
		"total_asp_blocks":  as.TotalASPBlocks,
		"total_html_blocks": as.TotalHTMLBlocks,
		"total_errors":      as.TotalErrors,
		"average_size":      as.AverageSize,
	}
}
