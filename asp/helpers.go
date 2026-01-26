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
	"strings"
)

// Example de uso do ASP Parser
/*
func main() {
	aspCode := `
<%
	Dim message
	message = "Hello from ASP"
	Response.Write(message)
%>
<html>
	<body>
		<h1>Welcome</h1>
	</body>
</html>
<%
	Set conn = CreateObject("ADODB.Connection")
%>
`

	parser := NewASPParser(aspCode)
	result, err := parser.Parse()

	if err != nil {
		fmt.Println("Error:", err)
		return
	}

	fmt.Printf("Total blocks: %d\n", len(result.Blocks))
	fmt.Printf("VB Programs: %d\n", len(result.VBPrograms))
	fmt.Printf("Parse errors: %d\n", len(result.Errors))

	for i, block := range result.Blocks {
		fmt.Printf("Block %d: %s (Line %d)\n", i, block.Type, block.Line)
		if block.Type == "asp" {
			fmt.Printf("  Content: %s\n", strings.TrimSpace(block.Content))
		}
	}

	for i, errors := range result.Errors {
		fmt.Printf("Error %d: %s\n", i, errors)
	}
}
*/

// PrintASPResult imprime informações sobre o resultado do parse
func PrintASPResult(result *ASPParserResult) {
	fmt.Println("=== ASP Parse Result ===")
	fmt.Printf("Total blocks: %d\n", len(result.Blocks))
	fmt.Printf("VB Programs: %d\n", len(result.VBPrograms))
	fmt.Printf("Parse errors: %d\n", len(result.Errors))
	fmt.Println()

	for i, block := range result.Blocks {
		fmt.Printf("[Block %d] Type: %s | Line: %d | Column: %d\n",
			i, block.Type, block.Line, block.Column)
		content := block.Content
		if len(content) > 50 {
			content = content[:50] + "..."
		}
		content = strings.ReplaceAll(content, "\n", "\\n")
		fmt.Printf("  Content: %s\n", content)
	}

	if len(result.Errors) > 0 {
		fmt.Println()
		fmt.Println("=== Errors ===")
		for i, err := range result.Errors {
			fmt.Printf("[Error %d] %v\n", i, err)
		}
	}
}

// ExtractHTMLOnly extrai apenas o conteúdo HTML de um arquivo ASP
func ExtractHTMLOnly(aspCode string) string {
	lexer := NewASPLexer(aspCode)
	blocks := lexer.Tokenize()

	var htmlParts []string
	for _, block := range blocks {
		if block.Type == "html" {
			htmlParts = append(htmlParts, block.Content)
		}
	}

	return strings.Join(htmlParts, "")
}

// ExtractVBScriptOnly extrai apenas o código VBScript
func ExtractVBScriptOnly(aspCode string) string {
	parser := NewASPParser(aspCode)
	return parser.ExtractVBScriptCode("\n\n")
}

// ValidateASP valida a sintaxe do código ASP
func ValidateASP(aspCode string) (bool, []error) {
	parser := NewASPParser(aspCode)
	result, err := parser.Parse()

	if err != nil {
		return false, []error{err}
	}

	if len(result.Errors) > 0 {
		return false, result.Errors
	}

	return true, nil
}
