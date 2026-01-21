package vbscript

import (
	"fmt"
	"testing"
)

func TestSimpleResponseEnd(t *testing.T) {
	tests := []struct {
		name string
		code string
	}{
		{"simple_member", "response.end"},
		{"with_newline", "response.end\n"},
		{"two_statements_colon", "x=1 : response.end"},
		{"in_case", "select case 1\ncase 1\nresponse.end\nend select"},
		{"in_case_colon", "select case 1\ncase 1\nx=1 : response.end\nend select"},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			defer func() {
				if r := recover(); r != nil {
					fmt.Printf("Code: %s\n", tt.code)
					t.Errorf("Parsing panicked: %v", r)
				}
			}()

			parser := NewParser(tt.code)
			program := parser.Parse()

			if program == nil {
				t.Errorf("Parsing returned nil program")
			}
		})
	}
}
