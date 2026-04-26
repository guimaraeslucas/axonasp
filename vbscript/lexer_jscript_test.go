package vbscript

import "testing"

func TestLexerEmitsASPJScriptBlockToken(t *testing.T) {
	lex := NewLexer(`<script runat="server" language="JScript">Response.Write("Hello")</script>`)
	lex.Mode = ModeASP
	lex.InASPBlock = false
	tok := lex.NextToken()
	block, ok := tok.(*ASPJScriptBlockToken)
	if !ok {
		t.Fatalf("expected ASPJScriptBlockToken, got %T", tok)
	}
	if block.Content != `Response.Write("Hello")` {
		t.Fatalf("unexpected block content: %q", block.Content)
	}
}

func TestLexerEmitsASPJScriptBlockTokenForSupportedScriptTagVariations(t *testing.T) {
	tags := []string{
		`<script type="text/javascript" language="javascript" runat="server">Response.Write("A")</script>`,
		`<script language="javascript" runat="server">Response.Write("B")</script>`,
		`<script type="text/javascript" language="jscript" runat="server">Response.Write("C")</script>`,
		`<script language="jscript" runat="server">Response.Write("D")</script>`,
	}

	for i := range tags {
		lex := NewLexer(tags[i])
		lex.Mode = ModeASP
		lex.InASPBlock = false
		tok := lex.NextToken()
		if _, ok := tok.(*ASPJScriptBlockToken); !ok {
			t.Fatalf("tag #%d: expected ASPJScriptBlockToken, got %T", i+1, tok)
		}
	}
}

func TestLexerRoutesPercentBlockToJScriptWhenPageDirectiveLanguageIsJScript(t *testing.T) {
	lex := NewLexer(`<%@ Language="JScript" %><% Response.Write(1 === 1 ? "ok" : "bad"); %>`)
	lex.Mode = ModeASP
	lex.InASPBlock = false

	if _, ok := lex.NextToken().(*ASPDirectiveStartToken); !ok {
		t.Fatalf("expected ASPDirectiveStartToken as first token")
	}

	for {
		tok := lex.NextToken()
		switch tok.(type) {
		case *ASPCodeEndToken:
			goto afterDirective
		case *EOFToken:
			t.Fatalf("unexpected EOF before directive end")
		}
	}

afterDirective:
	tok := lex.NextToken()
	block, ok := tok.(*ASPJScriptBlockToken)
	if !ok {
		t.Fatalf("expected ASPJScriptBlockToken after JScript directive, got %T", tok)
	}
	if block.Content == "" {
		t.Fatalf("expected non-empty jscript block content")
	}
}

func TestLexerRoutesCompactDirectiveThenPercentBlockToJScript(t *testing.T) {
	lex := NewLexer("<%@Language=\"JScript\"%>\r\n<% var metodo = String(\"POST\"); %>")
	lex.Mode = ModeASP
	lex.InASPBlock = false

	if _, ok := lex.NextToken().(*ASPDirectiveStartToken); !ok {
		t.Fatalf("expected ASPDirectiveStartToken as first token")
	}

	for {
		tok := lex.NextToken()
		switch tok.(type) {
		case *ASPCodeEndToken:
			goto afterDirective
		case *EOFToken:
			t.Fatalf("unexpected EOF before directive end")
		}
	}

afterDirective:
	tok := lex.NextToken()
	if _, ok := tok.(*ASPJScriptBlockToken); !ok {
		t.Fatalf("expected ASPJScriptBlockToken after compact JScript directive, got %T", tok)
	}
}

func TestLexerRoutesJScriptExpressionTagAsJScriptBlock(t *testing.T) {
	lex := NewLexer(`<%@ Language="JScript" %><%= "ok" %>`)
	lex.Mode = ModeASP
	lex.InASPBlock = false

	if _, ok := lex.NextToken().(*ASPDirectiveStartToken); !ok {
		t.Fatalf("expected ASPDirectiveStartToken as first token")
	}

	for {
		tok := lex.NextToken()
		switch tok.(type) {
		case *ASPCodeEndToken:
			goto afterDirective
		case *EOFToken:
			t.Fatalf("unexpected EOF before directive end")
		}
	}

afterDirective:
	tok := lex.NextToken()
	block, ok := tok.(*ASPJScriptBlockToken)
	if !ok {
		t.Fatalf("expected ASPJScriptBlockToken for JScript expression tag, got %T", tok)
	}
	if block.Content != `Response.Write("ok");` {
		t.Fatalf("unexpected converted expression content: %q", block.Content)
	}
}
