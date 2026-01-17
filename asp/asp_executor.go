package asp

import (
	"fmt"
	"strings"
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
		if block.Type == "html" {
			result.WriteString(block.Content)
		} else if block.Type == "asp" {
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
