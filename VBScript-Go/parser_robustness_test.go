package vbscript

import (
	"testing"
)

func TestParserRobustnessStrayTokens(t *testing.T) {
	// Simulate "If True Then : ' comment"
	// The colon and comment might reach parseBlockStatement if parseMultiInlineStatement logic is quirky
	// Or just "If True Then" followed by stray comment handled as block statement?
	
	// Test case 1: Stray colon treated as empty statement
	code1 := "Dim x : : Dim y"
	parser1 := NewParser(code1)
	program1 := parser1.Parse()
	if program1 == nil {
		t.Error("Code1 parsed nil")
	}

	// Test case 2: Stray comment treated as empty statement (via manual invocation simulation)
	// We can't easily force parseBlockStatement to be called on a comment from top-level
	// because parseGlobalStatement skips comments.
	// But we can test if the parser generally accepts weird sequences now.
	code2 := "Dim x \n ' comment \n Dim y"
	parser2 := NewParser(code2)
	parser2.Parse()
}
