package vbscript

import (
	"testing"

	"g3pix.com.br/axonasp/vbscript/ast"
)

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

func TestASPParser_AllowsUnquotedDirectiveValues(t *testing.T) {
	parser := NewASPParser(`<%@ Language=VBScript CODEPAGE="65001" EnableSessionState=False %>`)
	program := parser.Parse()
	if len(program.Body) != 1 {
		t.Fatalf("expected one statement, got %d", len(program.Body))
	}

	directive, ok := program.Body[0].(*ast.ASPDirectiveStatement)
	if !ok {
		t.Fatalf("expected ASP directive statement, got %T", program.Body[0])
	}
	if directive.Attributes["Language"] != "VBScript" {
		t.Fatalf("unexpected language directive value: %q", directive.Attributes["Language"])
	}
	if directive.Attributes["CODEPAGE"] != "65001" {
		t.Fatalf("unexpected codepage directive value: %q", directive.Attributes["CODEPAGE"])
	}
	if directive.Attributes["EnableSessionState"] != "False" {
		t.Fatalf("unexpected session directive value: %q", directive.Attributes["EnableSessionState"])
	}
}

func TestParser_AllowsLineContinuationWithUnderscore(t *testing.T) {
	code := `
Dim welcomeMessage
welcomeMessage = "Welcome to the enhanced VBScript parser! " + _
                 "This version supports colon separators, " + _
                 "plus operators for concatenation, " + _
                 "and line continuation with underscore."
`
	parser := NewParser(code)
	defer func() {
		if r := recover(); r != nil {
			t.Fatalf("parse panic: %v", r)
		}
	}()

	_ = parser.Parse()
}

func TestLexer_StripsBracketEscapedIdentifierName(t *testing.T) {
	lexer := NewLexer("[isEmpty]")
	token := lexer.NextToken()

	identifier, ok := token.(*IdentifierToken)
	if !ok {
		t.Fatalf("expected identifier token, got %T", token)
	}
	if identifier.Name != "isEmpty" {
		t.Fatalf("unexpected bracket-escaped identifier name: %q", identifier.Name)
	}
}

func TestLexer_BracketEscapedIdentifierAllowsSpaces(t *testing.T) {
	lexer := NewLexer("[My Variable]")
	token := lexer.NextToken()

	identifier, ok := token.(*IdentifierToken)
	if !ok {
		t.Fatalf("expected identifier token, got %T", token)
	}
	if identifier.Name != "My Variable" {
		t.Fatalf("unexpected bracket-escaped identifier name: %q", identifier.Name)
	}
}

func TestLexer_BracketEscapedIdentifierUnterminatedPanics(t *testing.T) {
	defer func() {
		if r := recover(); r == nil {
			t.Fatal("expected panic for unterminated bracket-escaped identifier")
		}
	}()

	lexer := NewLexer("[isEmpty")
	_ = lexer.NextToken()
}

func TestLexer_ASPStripsLeadingCRLFAfterPercentBlockEnd(t *testing.T) {
	lexer := NewLexer("<% x = 1 %>\r\nB")
	lexer.Mode = ModeASP

	for {
		tok := lexer.NextToken()
		if htmlTok, ok := tok.(*HTMLToken); ok {
			if htmlTok.Content != "B" {
				t.Fatalf("expected HTML content 'B', got %q", htmlTok.Content)
			}
			return
		}
		if _, ok := tok.(*EOFToken); ok {
			break
		}
	}

	t.Fatal("expected HTML token after ASP block end")
}

func TestLexer_ASPStripsLeadingLFAfterPercentBlockEnd(t *testing.T) {
	lexer := NewLexer("<% x = 1 %>\nB")
	lexer.Mode = ModeASP

	for {
		tok := lexer.NextToken()
		if htmlTok, ok := tok.(*HTMLToken); ok {
			if htmlTok.Content != "B" {
				t.Fatalf("expected HTML content 'B', got %q", htmlTok.Content)
			}
			return
		}
		if _, ok := tok.(*EOFToken); ok {
			break
		}
	}

	t.Fatal("expected HTML token after ASP block end")
}

func TestLexer_ASPStripsLeadingWhitespaceThenCRLFAfterPercentBlockEnd(t *testing.T) {
	lexer := NewLexer("<% x = 1 %>   \t\r\nB")
	lexer.Mode = ModeASP

	for {
		tok := lexer.NextToken()
		if htmlTok, ok := tok.(*HTMLToken); ok {
			if htmlTok.Content != "B" {
				t.Fatalf("expected HTML content 'B', got %q", htmlTok.Content)
			}
			return
		}
		if _, ok := tok.(*EOFToken); ok {
			break
		}
	}

	t.Fatal("expected HTML token after ASP block end")
}

func TestLexer_ASPStripsLeadingWhitespaceThenLFAfterScriptEnd(t *testing.T) {
	lexer := NewLexer("<script runat=\"server\">x = 1</script> \t\nB")
	lexer.Mode = ModeASP

	for {
		tok := lexer.NextToken()
		if htmlTok, ok := tok.(*HTMLToken); ok {
			if htmlTok.Content != "B" {
				t.Fatalf("expected HTML content 'B', got %q", htmlTok.Content)
			}
			return
		}
		if _, ok := tok.(*EOFToken); ok {
			break
		}
	}

	t.Fatal("expected HTML token after script block end")
}

func TestLexer_ASPPreservesLeadingWhitespaceWithoutLineBreakAfterPercentBlockEnd(t *testing.T) {
	lexer := NewLexer("<% x = 1 %>   B")
	lexer.Mode = ModeASP

	for {
		tok := lexer.NextToken()
		if htmlTok, ok := tok.(*HTMLToken); ok {
			if htmlTok.Content != "   B" {
				t.Fatalf("expected HTML content '   B', got %q", htmlTok.Content)
			}
			return
		}
		if _, ok := tok.(*EOFToken); ok {
			break
		}
	}

	t.Fatal("expected HTML token after ASP block end")
}

func TestLexer_ASPIncludeStripsLeadingWhitespaceThenCRLFAfterDirective(t *testing.T) {
	lexer := NewLexer("<!--#include file=\"header.inc\"-->   \t\r\nB")
	lexer.Mode = ModeASP

	for {
		tok := lexer.NextToken()
		if htmlTok, ok := tok.(*HTMLToken); ok {
			if htmlTok.Content != "B" {
				t.Fatalf("expected HTML content 'B', got %q", htmlTok.Content)
			}
			return
		}
		if _, ok := tok.(*EOFToken); ok {
			break
		}
	}

	t.Fatal("expected HTML token after include directive")
}

func TestLexer_ASPIncludePreservesWhitespaceWithoutLineBreakAfterDirective(t *testing.T) {
	lexer := NewLexer("<!--#include file=\"header.inc\"-->   B")
	lexer.Mode = ModeASP

	for {
		tok := lexer.NextToken()
		if htmlTok, ok := tok.(*HTMLToken); ok {
			if htmlTok.Content != "   B" {
				t.Fatalf("expected HTML content '   B', got %q", htmlTok.Content)
			}
			return
		}
		if _, ok := tok.(*EOFToken); ok {
			break
		}
	}

	t.Fatal("expected HTML token after include directive")
}

func TestLexer_ASPSuppressesFormattingHTMLBeforeCodeBlock(t *testing.T) {
	lexer := NewLexer("\r\n\t<% x = 1 %>\r\nB")
	lexer.Mode = ModeASP

	for {
		tok := lexer.NextToken()
		if htmlTok, ok := tok.(*HTMLToken); ok {
			if htmlTok.Content != "B" {
				t.Fatalf("expected HTML content 'B', got %q", htmlTok.Content)
			}
			return
		}
		if _, ok := tok.(*EOFToken); ok {
			break
		}
	}

	t.Fatal("expected HTML token after ASP block")
}

func TestLexer_ASPPreservesInlineSpaceBeforeCodeBlockWithoutLineBreak(t *testing.T) {
	lexer := NewLexer("A <% x = 1 %>B")
	lexer.Mode = ModeASP

	first := lexer.NextToken()
	htmlTok, ok := first.(*HTMLToken)
	if !ok {
		t.Fatalf("expected first token HTML, got %T", first)
	}
	if htmlTok.Content != "A " {
		t.Fatalf("expected HTML content 'A ', got %q", htmlTok.Content)
	}
}

func TestLexer_ASPCollapsesWhitespaceBetweenConsecutiveServerBlocks(t *testing.T) {
	lexer := NewLexer("<% x = 1 %>\r\n\r\n<% y = 2 %>\r\nB")
	lexer.Mode = ModeASP

	for {
		tok := lexer.NextToken()
		if htmlTok, ok := tok.(*HTMLToken); ok {
			if htmlTok.Content != "B" {
				t.Fatalf("expected HTML content 'B', got %q", htmlTok.Content)
			}
			return
		}
		if _, ok := tok.(*EOFToken); ok {
			break
		}
	}

	t.Fatal("expected HTML token after consecutive ASP blocks")
}
