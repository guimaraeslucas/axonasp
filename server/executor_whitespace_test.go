package server

import (
	"testing"

	"g3pix.com.br/axonasp/asp"
)

func TestShouldSuppressStructuralHTMLBlock_BetweenASPBlocks(t *testing.T) {
	blocks := []*asp.CodeBlock{
		{Type: "asp", Content: "Dim firstValue"},
		{Type: "html", Content: "\r\n\r\n"},
		{Type: "asp", Content: "Dim secondValue"},
	}

	if !shouldSuppressStructuralHTMLBlock(blocks, 1) {
		t.Fatalf("expected structural whitespace block to be suppressed")
	}
}

func TestShouldSuppressStructuralHTMLBlock_RealHTML(t *testing.T) {
	blocks := []*asp.CodeBlock{
		{Type: "asp", Content: "Dim firstValue"},
		{Type: "html", Content: "<div>\n</div>"},
		{Type: "asp", Content: "Dim secondValue"},
	}

	if shouldSuppressStructuralHTMLBlock(blocks, 1) {
		t.Fatalf("expected real HTML block not to be suppressed")
	}
}

func TestNormalizeStructuralBoundaryHTMLBlock_TrimsLeadingCRLF(t *testing.T) {
	blocks := []*asp.CodeBlock{
		{Type: "asp", Content: "Dim firstValue"},
		{Type: "html", Content: "\r\n\r\n<root/>"},
	}

	normalized := normalizeStructuralBoundaryHTMLBlock(blocks, 1, blocks[1].Content)
	if normalized != "<root/>" {
		t.Fatalf("expected leading boundary CRLF to be trimmed, got %q", normalized)
	}
}

func TestNormalizeStructuralBoundaryHTMLBlock_TrimsTrailingCRLF(t *testing.T) {
	blocks := []*asp.CodeBlock{
		{Type: "html", Content: "<root/>\r\n\r\n"},
		{Type: "asp", Content: "Dim firstValue"},
	}

	normalized := normalizeStructuralBoundaryHTMLBlock(blocks, 0, blocks[0].Content)
	if normalized != "<root/>" {
		t.Fatalf("expected trailing boundary CRLF to be trimmed, got %q", normalized)
	}
}
