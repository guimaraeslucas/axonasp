package asp

import (
	"strings"
	"testing"
)

func TestDateLiteralsExecution(t *testing.T) {
	code := `
<%
Dim d
d = #1/19/2026#
Response.Write TypeName(d) & "|" & Year(d)
%>`

	executor := NewASPExecutor()
	output, err := executor.Execute(code)
	if err != nil {
		t.Fatalf("Execution error: %v", err)
	}

	if !strings.Contains(output, "Date|2026") {
		t.Errorf("Expected 'Date|2026', got '%s'", output)
	}
}
