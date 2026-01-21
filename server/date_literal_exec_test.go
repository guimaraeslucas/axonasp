package server

import (
	"net/http/httptest"
	"testing"
	"time"

	vb "github.com/guimaraeslucas/vbscript-go"
)

func TestDateLiteralExecution(t *testing.T) {
	code := `
Dim d1, d2
d1 = #1/19/2026#
d2 = #2026-01-19#
t1 = TypeName(d1)
y1 = Year(d1)
`
	parser := vb.NewParser(code)
	program := parser.Parse()
	
	if program == nil {
		t.Fatal("Parsed program is nil")
	}

	rec := httptest.NewRecorder()
	req := httptest.NewRequest("GET", "/", nil)

	executor := &ASPExecutor{config: &ASPProcessorConfig{RootDir: "./www", ScriptTimeout: 30}}
	executor.context = NewExecutionContext(rec, req, "test-session", 5*time.Second)

	if err := executor.executeVBProgram(program); err != nil {
		t.Fatalf("execution failed: %v", err)
	}

	// Check d1
	val, ok := executor.context.GetVariable("d1")
	if !ok {
		t.Fatalf("d1 not found")
	}
	t.Logf("d1 type: %T, value: %v", val, val)
	
	// Check t1 (TypeName)
	t1, ok := executor.context.GetVariable("t1")
	if !ok {
		t.Fatal("t1 not found")
	}
	if t1 != "Date" {
		t.Errorf("Expected TypeName(d1)='Date', got '%v'", t1)
	}

	// Check y1 (Year)
	y1, ok := executor.context.GetVariable("y1")
	if !ok {
		t.Fatal("y1 not found")
	}
	if toInt(y1) != 2026 {
		t.Errorf("Expected Year(d1)=2026, got %v", y1)
	}
}
