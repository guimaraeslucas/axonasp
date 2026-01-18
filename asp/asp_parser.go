package asp

import (
	"fmt"
	"strings"

	vb "github.com/guimaraeslucas/vbscript-go"
	"github.com/guimaraeslucas/vbscript-go/ast"
)

// ASPParserResult contém o resultado da análise de código ASP
type ASPParserResult struct {
	Blocks      []*CodeBlock         // Todos os blocos de código
	VBPrograms  map[int]*ast.Program // Programas VBScript por índice do bloco
	Errors      []error              // Erros durante análise
	HTMLContent []string             // Conteúdo HTML em ordem
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
	vbBlockIndex := 0
	for i, block := range blocks {
		switch block.Type {
		case "directive":
			// ASP directives like <%@ Language=VBScript %>
			// These are processed but don't generate code
			// They are used for configuration (Language, CodePage, etc.)
			// No action needed here - directives are parsed in the lexer
		case "asp":
			// Tenta fazer parse do bloco VBScript
			program, err := ap.parseVBBlock(block.Content)
			if err != nil {
				parseErr := fmt.Errorf("Error: %d: %v", block.Line, err)
				result.Errors = append(result.Errors, parseErr)
				if ap.options.DebugMode {
					fmt.Printf("[ASP Parser Error] Line %d: %v\n", block.Line, err)
				}
			}
			result.VBPrograms[i] = program
			vbBlockIndex++
		case "html":
			// Armazena conteúdo HTML
			result.HTMLContent = append(result.HTMLContent, block.Content)
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
				Body:           []ast.Statement{},
			},
			nil
	}

	// Pre-process colons to handle multi-statement lines
	// The VBScript parser might panic on colons, so we convert them to newlines
	processedCode := preProcessColons(trimmedCode)

	// Usa o parser do VBScript-Go
	parser := vb.NewParserWithOptions(processedCode, ap.vbOptions)

	// Faz o parse e captura possíveis panics
	defer func() {
		if r := recover(); r != nil {
			err = fmt.Errorf("panic durante parse: %v", r)
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
	
	for i := 0; i < len(code); i++ {
		char := code[i]
		
		if char == '\n' || char == '\r' {
			inComment = false
			inString = false // Strings don't span lines in VBScript
			sb.WriteByte(char)
			continue
		}

		if inComment {
			sb.WriteByte(char)
			continue
		}

		if char == '"' {
			inString = !inString
			sb.WriteByte(char)
			continue
		}
		
		if inString {
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
				sb.WriteString(code[i:i+4])
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
