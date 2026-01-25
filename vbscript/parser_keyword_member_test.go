package vbscript

import (
	"testing"
)

func TestKeywordAsMemberName(t *testing.T) {
	tests := []struct {
		name string
		code string
	}{
		{"end", "Response.end"},
		{"if", "obj.if"},
		{"then", "obj.then"},
		{"else", "obj.else"},
		{"for", "obj.for"},
		{"next", "obj.next"},
		{"select", "obj.select"},
		{"case", "obj.case"},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			defer func() {
				if r := recover(); r != nil {
					t.Errorf("Parsing '%s' panicked: %v", tt.code, r)
				}
			}()

			parser := NewParser(tt.code)
			program := parser.Parse()

			if program == nil {
				t.Errorf("Parsing '%s' returned nil program", tt.code)
			}
		})
	}
}
