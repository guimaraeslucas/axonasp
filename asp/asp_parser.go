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
}

// NewASPParsingOptions cria opções padrão
func NewASPParsingOptions() *ASPParsingOptions {
	return &ASPParsingOptions{
		SaveComments:      false,
		StrictMode:        false,
		AllowImplicitVars: true,
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
				result.Errors = append(result.Errors, fmt.Errorf("erro no bloco ASP linha %d: %v", block.Line, err))
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

	// Usa o parser do VBScript-Go
	parser := vb.NewParserWithOptions(trimmedCode, ap.vbOptions)

	// Faz o parse e captura possíveis panics
	defer func() {
		if r := recover(); r != nil {
			err = fmt.Errorf("panic durante parse: %v", r)
		}
	}()

	program = parser.Parse()
	return program, nil
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