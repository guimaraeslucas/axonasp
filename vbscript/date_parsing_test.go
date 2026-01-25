package vbscript

import (
	"testing"
	"time"
)

func TestGetDate(t *testing.T) {
	tests := []struct {
		input    string
		expected time.Time
		hasError bool
	}{
		{"1/19/2026", time.Date(2026, 1, 19, 0, 0, 0, 0, time.UTC), false},
		{"2026-01-19", time.Date(2026, 1, 19, 0, 0, 0, 0, time.UTC), false},
		{"1/2/2006", time.Date(2006, 1, 2, 0, 0, 0, 0, time.UTC), false},
		{"01/02/2006", time.Date(2006, 1, 2, 0, 0, 0, 0, time.UTC), false},
		{"1/1/2023", time.Date(2023, 1, 1, 0, 0, 0, 0, time.UTC), false},
	}

	for _, test := range tests {
		result, err := GetDate(test.input)
		if test.hasError {
			if err == nil {
				t.Errorf("Expected error for input %s, but got none", test.input)
			}
		} else {
			if err != nil {
				t.Errorf("Unexpected error for input %s: %v", test.input, err)
			}
			// Compare year, month, day. Ignore timezone for this basic check or ensure UTC.
			// GetDate might return local time or UTC depending on Parse.
			// time.Parse returns UTC if no timezone info.
			if result.Year() != test.expected.Year() || result.Month() != test.expected.Month() || result.Day() != test.expected.Day() {
				t.Errorf("For input %s, expected %v, got %v", test.input, test.expected, result)
			}
		}
	}
}

func TestLexerDateLiteral(t *testing.T) {
    code := "#1/19/2026#"
    lexer := NewLexer(code)
    token := lexer.NextToken()
    
    dateToken, ok := token.(*DateLiteralToken)
    if !ok {
        t.Fatalf("Expected DateLiteralToken, got %T", token)
    }
    
    expected := time.Date(2026, 1, 19, 0, 0, 0, 0, time.UTC)
    if dateToken.Value.Year() != expected.Year() || dateToken.Value.Month() != expected.Month() || dateToken.Value.Day() != expected.Day() {
        t.Errorf("Expected date %v, got %v", expected, dateToken.Value)
    }
}

func TestLexerDateLiteralWithSpaces(t *testing.T) {
    code := "# 1/19/2026 #"
    lexer := NewLexer(code)
    token := lexer.NextToken()
    
    dateToken, ok := token.(*DateLiteralToken)
    if !ok {
        t.Fatalf("Expected DateLiteralToken, got %T", token)
    }
    
    expected := time.Date(2026, 1, 19, 0, 0, 0, 0, time.UTC)
    if dateToken.Value.Year() != expected.Year() || dateToken.Value.Month() != expected.Month() || dateToken.Value.Day() != expected.Day() {
        t.Errorf("Expected date %v, got %v", expected, dateToken.Value)
    }
}
