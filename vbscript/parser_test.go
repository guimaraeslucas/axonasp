package vbscript

import "testing"

func TestParser_AllowsEmptyLiteralAsArgument(t *testing.T) {
	code := `
Dim customConfig
Dim extraConfig
extraConfig = (new JSON)( empty, customConfig, false )
`
	parser := NewParser(code)
	defer func() {
		if r := recover(); r != nil {
			t.Fatalf("parse panic: %v", r)
		}
	}()

	_ = parser.Parse()
}

func TestParser_AllowsEmptyLiteralInCallWithoutParens(t *testing.T) {
	code := `
Dim editor
editor.replaceAll empty
`
	parser := NewParser(code)
	defer func() {
		if r := recover(); r != nil {
			t.Fatalf("parse panic: %v", r)
		}
	}()

	_ = parser.Parse()
}

func TestParser_AllowsNewDefaultCallInsideClassMethod(t *testing.T) {
	code := `
Class CKEditor
	Public Function ReplaceInstance(id)
		Dim customConfig
		Dim extraConfig
		extraConfig = (new JSON)( empty, customConfig, false )
	End Function
End Class
`
	parser := NewParser(code)
	defer func() {
		if r := recover(); r != nil {
			t.Fatalf("parse panic: %v", r)
		}
	}()

	_ = parser.Parse()
}
