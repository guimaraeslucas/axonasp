package vbscript

import (
	"testing"
)

func TestParserColonStandalone(t *testing.T) {
	code := ": ' comment"
	
	defer func() {
		if r := recover(); r != nil {
			t.Errorf("Parsing failed with panic: %v", r)
		}
	}()

	parser := NewParser(code)
	program := parser.Parse()
	
	if program == nil {
		t.Error("Parsed program is nil")
	}
}
