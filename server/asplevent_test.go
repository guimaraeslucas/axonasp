package server

import (
	"net/http/httptest"
	"strings"
	"testing"
)

func TestAsplEventHandling(t *testing.T) {
	code := `
		<!-- #include file="asplite/asplite.asp"-->
		<%
		dim asplEvent : asplEvent = aspL.getRequest("asplEvent")
		
		Response.Write "asplEvent = '" & asplEvent & "'" & vbCrLf
		Response.Write "isEmpty(asplEvent) = " & aspL.isEmpty(asplEvent) & vbCrLf
		Response.Write "lcase(asplEvent) = '" & lcase(asplEvent) & "'" & vbCrLf
		
		select case lcase(asplEvent)
			case "test1"
				Response.Write "Matched: test1"
			case "test2"
				Response.Write "Matched: test2"
			case else
				Response.Write "No match - default case"
		end select
		%>
	`

	w := httptest.NewRecorder()
	r := httptest.NewRequest("GET", "/test.asp?asplEvent=test1", nil)

	processor := NewASPProcessor(&ASPProcessorConfig{
		RootDir:       "../www",
		ScriptTimeout: 30,
		DebugASP:      false,
	})

	err := processor.ExecuteASPFile(code, "../www/test.asp", w, r)
	if err != nil {
		t.Fatalf("ExecuteASPFile failed: %v", err)
	}

	output := w.Body.String()
	t.Logf("Output:\n%s", output)

	if !strings.Contains(output, "asplEvent = 'test1'") {
		t.Errorf("Expected asplEvent to be 'test1', got: %s", output)
	}

	if !strings.Contains(output, "Matched: test1") {
		t.Errorf("Expected 'Matched: test1' in output, got: %s", output)
	}
}

func TestAsplEventInHandler(t *testing.T) {
	code := `
		<!-- #include file="asplite/asplite.asp"-->
		<%
		select case lcase(aspL.getRequest("asplEvent"))
			case "sampleform18"
				Response.Write "JSON for sampleform18"
			case else
				Response.Write "Default HTML page"
		end select
		%>
	`

	tests := []struct {
		name           string
		queryString    string
		expectedOutput string
	}{
		{"With asplEvent", "?asplEvent=sampleform18", "JSON for sampleform18"},
		{"Without asplEvent", "", "Default HTML page"},
	}

	processor := NewASPProcessor(&ASPProcessorConfig{
		RootDir:       "../www",
		ScriptTimeout: 30,
		DebugASP:      false,
	})

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			w := httptest.NewRecorder()
			r := httptest.NewRequest("GET", "/test.asp"+tt.queryString, nil)

			err := processor.ExecuteASPFile(code, "../www/test.asp", w, r)
			if err != nil {
				t.Fatalf("ExecuteASPFile failed: %v", err)
			}

			output := w.Body.String()
			t.Logf("Output:\n%s", output)

			if !strings.Contains(output, tt.expectedOutput) {
				t.Errorf("Expected '%s' in output, got: %s", tt.expectedOutput, output)
			}
		})
	}
}
