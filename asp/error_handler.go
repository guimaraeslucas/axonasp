package asp

// ASPError representa um erro específico do ASP
type ASPError struct {
	Code       string // Código de erro
	Message    string // Mensagem de erro
	Line       int    // Linha onde ocorreu o erro
	Column     int    // Coluna onde ocorreu o erro
	BlockIndex int    // Índice do bloco ASP
	Context    string // Contexto do erro (parte do código)
}

// Error implementa a interface error
func (ae *ASPError) Error() string {
	if ae.Line > 0 && ae.Column > 0 {
		return ae.Code + ": " + ae.Message + " (Linha " + string(rune(ae.Line)) + ", Coluna " + string(rune(ae.Column)) + ")"
	}
	return ae.Code + ": " + ae.Message
}

// Códigos de erro ASP comuns
const (
	ERR_SYNTAX             = "ASP0001"
	ERR_INVALID_BLOCK      = "ASP0002"
	ERR_UNCLOSED_BLOCK     = "ASP0003"
	ERR_PARSER_ERROR       = "ASP0004"
	ERR_INVALID_OBJECT     = "ASP0005"
	ERR_METHOD_NOT_FOUND   = "ASP0006"
	ERR_PROPERTY_NOT_FOUND = "ASP0007"
)

// NewASPError cria um novo erro ASP
func NewASPError(code, message string, line, column int) *ASPError {
	return &ASPError{
		Code:    code,
		Message: message,
		Line:    line,
		Column:  column,
	}
}

// ASPErrorHandler trata erros ASP
type ASPErrorHandler struct {
	errors []*ASPError
}

// NewASPErrorHandler cria um novo manipulador de erros
func NewASPErrorHandler() *ASPErrorHandler {
	return &ASPErrorHandler{
		errors: make([]*ASPError, 0),
	}
}

// AddError adiciona um erro
func (aeh *ASPErrorHandler) AddError(err *ASPError) {
	aeh.errors = append(aeh.errors, err)
}

// GetErrors retorna todos os erros
func (aeh *ASPErrorHandler) GetErrors() []*ASPError {
	return aeh.errors
}

// HasErrors retorna true se houver erros
func (aeh *ASPErrorHandler) HasErrors() bool {
	return len(aeh.errors) > 0
}

// Clear limpa a lista de erros
func (aeh *ASPErrorHandler) Clear() {
	aeh.errors = make([]*ASPError, 0)
}

// GetErrorCount retorna o número de erros
func (aeh *ASPErrorHandler) GetErrorCount() int {
	return len(aeh.errors)
}
