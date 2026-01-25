package vbscript

import (
	"testing"
)

func TestResponseEndInSelectCase(t *testing.T) {
	code := `select case "test"
	case "php"
		response.Status="404" : response.end			
end select`

	defer func() {
		if r := recover(); r != nil {
			t.Errorf("Parsing Response.end in select case panicked: %v", r)
		}
	}()

	parser := NewParser(code)
	program := parser.Parse()

	if program == nil {
		t.Errorf("Parsing returned nil program")
	}
}
