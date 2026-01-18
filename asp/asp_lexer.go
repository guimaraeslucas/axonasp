package asp

import (
	"strings"
)

// CodeBlock representa um bloco de código ASP ou conteúdo HTML
type CodeBlock struct {
	Type       string // "html", "asp", "directive", "text"
	Content    string
	Line       int
	Column     int
	StartPos   int
	EndPos     int
	Attributes map[string]string // For directives (e.g., Language=VBScript)
}

// ASPLexer realiza análise léxica de código ASP clássico
// Ele identifica blocos entre <% %> como código VBScript e ignora o resto
type ASPLexer struct {
	Code          string
	Index         int
	CurrentLine   int
	CurrentColumn int
	Length        int
	blocks        []*CodeBlock
}

// NewASPLexer cria uma nova instância de ASPLexer
func NewASPLexer(code string) *ASPLexer {
	return &ASPLexer{
		Code:          code,
		Index:         0,
		CurrentLine:   1,
		CurrentColumn: 0,
		Length:        len(code),
		blocks:        make([]*CodeBlock, 0),
	}
}

// Tokenize realiza a análise léxica do código ASP
// Retorna uma lista de blocos de código identificados
func (al *ASPLexer) Tokenize() []*CodeBlock {
	al.blocks = make([]*CodeBlock, 0)

	for al.Index < al.Length {
		// Procura pelo início de um bloco ASP
		aspStart := al.findNextASPBlock()

		if aspStart == -1 {
			// Não há mais blocos ASP, adiciona o conteúdo restante como HTML
			if al.Index < al.Length {
				content := al.Code[al.Index:]
				if strings.TrimSpace(content) != "" {
					al.blocks = append(al.blocks, &CodeBlock{
						Type:     "html",
						Content:  content,
						Line:     al.CurrentLine,
						Column:   al.CurrentColumn,
						StartPos: al.Index,
						EndPos:   al.Length,
					})
				}
			}
			break
		}

		// Adiciona conteúdo HTML anterior ao bloco ASP
		if aspStart > al.Index {
			htmlContent := al.Code[al.Index:aspStart]
			al.blocks = append(al.blocks, &CodeBlock{
				Type:     "html",
				Content:  htmlContent,
				Line:     al.CurrentLine,
				Column:   al.CurrentColumn,
				StartPos: al.Index,
				EndPos:   aspStart,
			})
			al.updatePosition(htmlContent)
		}

		// Processa o bloco ASP ou diretiva
		if al.isDirective(aspStart) {
			al.processDirective(aspStart)
		} else {
			al.processASPBlock(aspStart)
		}
	}

	return al.blocks
}

// findNextASPBlock encontra a próxima ocorrência de <% ou <%@
func (al *ASPLexer) findNextASPBlock() int {
	search := al.Code[al.Index:]
	idx := strings.Index(search, "<%")

	if idx == -1 {
		return -1
	}

	return al.Index + idx
}

// isDirective checks if the block starting at position is a directive (<%@ ... %>)
func (al *ASPLexer) isDirective(startPos int) bool {
	if startPos+3 >= al.Length {
		return false
	}
	return al.Code[startPos:startPos+3] == "<%@"
}

// findASPBlockEnd encontra a próxima ocorrência de %>
func (al *ASPLexer) findASPBlockEnd(startPos int) int {
	search := al.Code[startPos:]
	idx := strings.Index(search, "%>")

	if idx == -1 {
		return -1
	}

	return startPos + idx + 2 // +2 para incluir %>
}

// processASPBlock extrai e processa um bloco de código ASP
func (al *ASPLexer) processASPBlock(startPos int) {
	blockStart := startPos + 2 // Pula <%
	blockEnd := al.findASPBlockEnd(blockStart)

	if blockEnd == -1 {
		// Bloco não foi fechado corretamente, trata como HTML e avança para evitar loop infinito
		htmlContent := al.Code[startPos:]
		al.blocks = append(al.blocks, &CodeBlock{
			Type:     "html",
			Content:  htmlContent,
			Line:     al.CurrentLine,
			Column:   al.CurrentColumn,
			StartPos: startPos,
			EndPos:   al.Length,
		})
		al.Index = al.Length // Avança até o fim
		return
	}

	// Check for <%= (Response.Write shorthand)
	isOutput := false
	if startPos+3 <= al.Length && al.Code[startPos:startPos+3] == "<%=" {
		isOutput = true
		blockStart++ // Skip =
	}

	// Extrai o conteúdo do bloco ASP (sem %> no final)
	// Remove espaçamento em branco do início e fim, compatível com ASP clássico
	content := strings.TrimSpace(al.Code[blockStart : blockEnd-2])

	if isOutput {
		content = "Response.Write(" + content + ")"
	}

	al.blocks = append(al.blocks, &CodeBlock{
		Type:     "asp",
		Content:  content,
		Line:     al.CurrentLine,
		Column:   al.CurrentColumn,
		StartPos: startPos,
		EndPos:   blockEnd,
	})

	// Atualiza posição
	processedContent := al.Code[al.Index:blockEnd]
	al.updatePosition(processedContent)
}

// updatePosition atualiza a linha e coluna atual baseado no conteúdo processado
func (al *ASPLexer) updatePosition(content string) {
	lines := strings.Split(content, "\n")

	if len(lines) > 1 {
		al.CurrentLine += len(lines) - 1
		al.CurrentColumn = len(lines[len(lines)-1])
	} else {
		al.CurrentColumn += len(content)
	}

	al.Index += len(content)
}

// GetAspBlocks retorna apenas os blocos de código ASP
func (al *ASPLexer) GetAspBlocks() []*CodeBlock {
	var aspBlocks []*CodeBlock
	for _, block := range al.blocks {
		if block.Type == "asp" {
			aspBlocks = append(aspBlocks, block)
		}
	}
	return aspBlocks
}

// GetAllBlocks retorna todos os blocos
func (al *ASPLexer) GetAllBlocks() []*CodeBlock {
	return al.blocks
}

// processDirective processes an ASP directive like <%@ Language=VBScript %>
func (al *ASPLexer) processDirective(startPos int) {
	blockStart := startPos + 3 // Skip <%@
	blockEnd := al.findASPBlockEnd(blockStart)

	if blockEnd == -1 {
		// Directive not closed properly, treat as HTML and advance to avoid infinite loop
		htmlContent := al.Code[startPos:]
		al.blocks = append(al.blocks, &CodeBlock{
			Type:     "html",
			Content:  htmlContent,
			Line:     al.CurrentLine,
			Column:   al.CurrentColumn,
			StartPos: startPos,
			EndPos:   al.Length,
		})
		al.Index = al.Length // Avança até o fim
		return
	}

	// Extract directive content (without %> at end)
	content := strings.TrimSpace(al.Code[blockStart : blockEnd-2])

	// Parse directive attributes
	attributes := al.parseDirectiveAttributes(content)

	al.blocks = append(al.blocks, &CodeBlock{
		Type:       "directive",
		Content:    content,
		Line:       al.CurrentLine,
		Column:     al.CurrentColumn,
		StartPos:   startPos,
		EndPos:     blockEnd,
		Attributes: attributes,
	})

	// Update position
	processedContent := al.Code[al.Index:blockEnd]
	al.updatePosition(processedContent)
}

// parseDirectiveAttributes parses attributes from directive content
// Example: "Language=VBScript" -> {"Language": "VBScript"}
func (al *ASPLexer) parseDirectiveAttributes(content string) map[string]string {
	attributes := make(map[string]string)

	// Split by whitespace to get individual attribute=value pairs
	parts := strings.Fields(content)

	for _, part := range parts {
		// Split by = to get key and value
		if idx := strings.Index(part, "="); idx != -1 {
			key := strings.TrimSpace(part[:idx])
			value := strings.TrimSpace(part[idx+1:])

			// Remove quotes if present
			value = strings.Trim(value, `"'`)

			attributes[key] = value
		}
	}

	return attributes
}

// Reset reinicia o lexer para o início
func (al *ASPLexer) Reset() {
	al.Index = 0
	al.CurrentLine = 1
	al.CurrentColumn = 0
	al.blocks = make([]*CodeBlock, 0)
}
