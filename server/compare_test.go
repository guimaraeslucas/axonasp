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

func TestForwardFunctionCallHoisted(t *testing.T) {
	code := "result = Answer()\nFunction Answer()\n    Answer = 42\nEnd Function\n"
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

	if toInt(val) != 42 {
		t.Fatalf("expected Answer() to run before declaration and yield 42, got %#v", val)
	}
}

func TestForwardCallAcrossBlocksAndClass(t *testing.T) {
	block1 := "Class C\nSub Run()\n    result = Outside()\nEnd Sub\nEnd Class\nDim o\nSet o = New C\no.Run\n"
	block2 := "Function Outside()\n    Outside = 99\nEnd Function\n"

	prog1 := vb.NewParser(block1).Parse()
	prog2 := vb.NewParser(block2).Parse()

	blocks := []*asp.CodeBlock{
		{Type: "asp", Content: block1, Line: 1},
		{Type: "asp", Content: block2, Line: 1},
	}

	result := &asp.ASPParserResult{
		Blocks:     blocks,
		VBPrograms: map[int]*ast.Program{0: prog1, 1: prog2},
	}

	rec := httptest.NewRecorder()
	req := httptest.NewRequest("GET", "/", nil)

	executor := &ASPExecutor{config: &ASPProcessorConfig{RootDir: "./www", ScriptTimeout: 30}}
	executor.context = NewExecutionContext(rec, req, "test-session", 5*time.Second)

	if err := executor.executeBlocks(result); err != nil {
		t.Fatalf("execution failed: %v", err)
	}

	val, ok := executor.context.GetVariable("result")
	if !ok {
		t.Fatalf("expected result variable to be set")
	}

	if toInt(val) != 99 {
		t.Fatalf("expected Outside() to be callable from class in earlier block, got %#v", val)
	}
}
