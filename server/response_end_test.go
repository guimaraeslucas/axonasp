package server

import (
	"net/http/httptest"
	"strings"
	"testing"
)

func TestResponseEnd(t *testing.T) {
	code := `
		<%
		Response.Write "Before Response.End"
		Response.End
		Response.Write "After Response.End - this should NOT appear"
		%>
	`

	w := httptest.NewRecorder()
	r := httptest.NewRequest("GET", "/test.asp", nil)

	processor := NewASPProcessor(&ASPProcessorConfig{
		RootDir:       "../www",
		ScriptTimeout: 30,
		DebugASP:      false,
	})

	err := processor.ExecuteASPFile(code, "../www/test.asp", w, r)

	// Response.End should NOT return an error - it's a control flow signal
	if err != nil {
		t.Logf("ExecuteASPFile returned error: %v (this may be expected)", err)
	}

	output := w.Body.String()
	t.Logf("Output: %q", output)

	if !strings.Contains(output, "Before Response.End") {
		t.Errorf("Expected 'Before Response.End' in output")
	}

	if strings.Contains(output, "After Response.End") {
		t.Errorf("Found 'After Response.End' in output - Response.End did not stop execution!")
	}
}

func TestResponseEndInSubroutine(t *testing.T) {
	code := `
		<%
		Sub TestSub()
			Response.Write "In subroutine before Response.End"
			Response.End
			Response.Write "In subroutine after Response.End - should NOT appear"
		End Sub
		
		Response.Write "Before calling TestSub"
		TestSub()
		Response.Write "After calling TestSub - should NOT appear"
		%>
	`

	w := httptest.NewRecorder()
	r := httptest.NewRequest("GET", "/test.asp", nil)

	processor := NewASPProcessor(&ASPProcessorConfig{
		RootDir:       "../www",
		ScriptTimeout: 30,
		DebugASP:      false,
	})

	err := processor.ExecuteASPFile(code, "../www/test.asp", w, r)

	if err != nil {
		t.Logf("ExecuteASPFile returned error: %v", err)
	}

	output := w.Body.String()
	t.Logf("Output: %q", output)

	if !strings.Contains(output, "Before calling TestSub") {
		t.Errorf("Expected 'Before calling TestSub' in output")
	}

	if !strings.Contains(output, "In subroutine before Response.End") {
		t.Errorf("Expected 'In subroutine before Response.End' in output")
	}

	if strings.Contains(output, "In subroutine after Response.End") {
		t.Errorf("Found 'In subroutine after Response.End' - Response.End did not stop execution!")
	}

	if strings.Contains(output, "After calling TestSub") {
		t.Errorf("Found 'After calling TestSub' - Response.End should have stopped all execution!")
	}
}

func TestResponseEndInExecuteGlobal(t *testing.T) {
	code := `
		<!-- #include file="asplite/asplite.asp"-->
		<%
		Response.Write "Before ExecuteGlobal"
		
		Dim codeToExecute
		codeToExecute = "Response.Write ""Inside ExecuteGlobal before Response.End"": Response.End: Response.Write ""Inside ExecuteGlobal after Response.End - should NOT appear"""
		
		ExecuteGlobal codeToExecute
		
		Response.Write "After ExecuteGlobal - should NOT appear"
		%>
	`

	w := httptest.NewRecorder()
	r := httptest.NewRequest("GET", "/test.asp", nil)

	processor := NewASPProcessor(&ASPProcessorConfig{
		RootDir:       "../www",
		ScriptTimeout: 30,
		DebugASP:      false,
	})

	err := processor.ExecuteASPFile(code, "../www/test.asp", w, r)

	if err != nil {
		t.Logf("ExecuteASPFile returned error: %v", err)
	}

	output := w.Body.String()
	t.Logf("Output: %q", output)

	if !strings.Contains(output, "Before ExecuteGlobal") {
		t.Errorf("Expected 'Before ExecuteGlobal' in output")
	}

	if !strings.Contains(output, "Inside ExecuteGlobal before Response.End") {
		t.Errorf("Expected 'Inside ExecuteGlobal before Response.End' in output")
	}

	if strings.Contains(output, "Inside ExecuteGlobal after Response.End") {
		t.Errorf("Found 'Inside ExecuteGlobal after Response.End' - Response.End did not stop execution!")
	}

	if strings.Contains(output, "After ExecuteGlobal") {
		t.Errorf("Found 'After ExecuteGlobal' - Response.End should have stopped all execution!")
	}
}
