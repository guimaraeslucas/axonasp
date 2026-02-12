package server

import (
	"net/http/httptest"
	"strings"
	"testing"
)

func TestFunctionCall_GlobalDeclarationNotShadowedByParentReturnVariable(t *testing.T) {
	aspCode := `<%
Function parent(flag)
    If flag = "call-child" Then
        parent = child()
    Else
        parent = flag
    End If
End Function

Function child()
    child = parent("ok")
End Function

Response.Write parent("call-child")
%>`

	w := httptest.NewRecorder()
	r := httptest.NewRequest("GET", "/test.asp", nil)

	executor := NewASPExecutor(&ASPProcessorConfig{
		RootDir:       ".",
		ScriptTimeout: 30,
	})
	err := executor.Execute(aspCode, "/test.asp", w, r, "test-session-function-shadow")
	if err != nil {
		t.Fatalf("Execute error: %v", err)
	}

	body := w.Body.String()
	if !strings.Contains(body, "ok") {
		t.Fatalf("expected output to contain %q, got %q", "ok", body)
	}
}
