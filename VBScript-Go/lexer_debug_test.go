package vbscript

import (
	"fmt"
	"testing"
)

func TestDebugColonParsing(t *testing.T) {
	code := "x=1 : response.end"
	
	lexer := NewLexer(code)
	tokens := []Token{}
	for {
		tok := lexer.NextToken()
		tokens = append(tokens, tok)
		fmt.Printf("Token %d: %T = %v\n", len(tokens)-1, tok, tok)
		if _, ok := tok.(*EOFToken); ok {
			break
		}
	}
}
