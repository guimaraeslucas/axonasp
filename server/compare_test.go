package server

import (
	"net/http/httptest"
	"testing"
	"time"

	vb "github.com/guimaraeslucas/vbscript-go"
	"github.com/guimaraeslucas/vbscript-go/ast"
)

func TestCompareEqualTextModeAccents(t *testing.T) {
	if !compareEqual("ÁRVORE", "árvore", ast.OptionCompareText) {
		t.Fatalf("expected text compare to ignore case for accented characters")
	}

	if compareEqual("CAFÉ", "cafe", ast.OptionCompareText) {
		t.Fatalf("text compare should still honor diacritics")
	}

	if compareEqual("CAFÉ", "café", ast.OptionCompareBinary) {
		t.Fatalf("binary compare should be case-sensitive")
	}
}

func TestCompareLessStringOperandsLexical(t *testing.T) {
	if !compareLess("10", "2", ast.OptionCompareBinary) {
		t.Fatalf("string operands should compare lexically, not numerically")
	}

	if compareLess("10", 2, ast.OptionCompareBinary) {
		t.Fatalf("mixed string and numeric should compare numerically")
	}
}

func TestOptionCompareAppliedDuringExecution(t *testing.T) {
	code := "Option Compare Text\nresult = \"ÁRVORE\" = \"árvore\"\n"
	parser := vb.NewParser(code)
	program := parser.Parse()

	rec := httptest.NewRecorder()
	req := httptest.NewRequest("GET", "/", nil)

	executor := &ASPExecutor{config: &ASPProcessorConfig{RootDir: "./www", ScriptTimeout: 30}}
	executor.context = NewExecutionContext(rec, req, "test-session", 5*time.Second)

	if err := executor.executeVBProgram(program); err != nil {
		t.Fatalf("execution failed: %v", err)
	}

	val, ok := executor.context.GetVariable("result")
	if !ok {
		t.Fatalf("expected result variable to be set")
	}

	b, ok := val.(bool)
	if !ok || !b {
		t.Fatalf("expected text comparison to yield true, got %#v", val)
	}
}
