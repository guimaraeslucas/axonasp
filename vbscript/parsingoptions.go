package vbscript

// ParsingOptions contains configuration options for parsing VBScript
type ParsingOptions struct {
	// SaveComments indicates whether comments should be preserved in the AST
	SaveComments bool
}

// NewParsingOptions creates a new ParsingOptions with default values
func NewParsingOptions() *ParsingOptions {
	return &ParsingOptions{
		SaveComments: false,
	}
}
