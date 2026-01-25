package server

import (
	"net/http/httptest"
	"path/filepath"
	"strings"
	"testing"
)

func TestDictionaryAccess(t *testing.T) {
	code := `
		<%
		' Test dictionary creation and access
		Dim dict
		Set dict = Server.CreateObject("Scripting.Dictionary")
		dict.Add "name", "field1"
		dict.Add "type", "text"
		
		Response.Write "Direct access - name: " & dict("name") & vbCrLf
		Response.Write "Direct access - type: " & dict("type") & vbCrLf
		Response.Write "TypeName: " & TypeName(dict) & vbCrLf
		Response.Write "Count: " & dict.Count & vbCrLf
		
		' Test storing in variable
		Dim dict2
		Set dict2 = dict
		Response.Write "Via variable - name: " & dict2("name") & vbCrLf
		Response.Write "Via variable - type: " & dict2("type") & vbCrLf
		
		' Test storing in array
		Dim arr : arr = Array()
		ReDim arr(0)
		Set arr(0) = dict
		
		Response.Write "From array - name: " & arr(0)("name") & vbCrLf
		Response.Write "From array - type: " & arr(0)("type") & vbCrLf
		
		' Test with For Each
		Dim item
		For Each item In arr
			Response.Write "ForEach - TypeName: " & TypeName(item) & vbCrLf
			Response.Write "ForEach - Count: " & item.Count & vbCrLf
			Response.Write "ForEach - name: " & item("name") & vbCrLf
			Response.Write "ForEach - type: " & item("type") & vbCrLf
		Next
		%>
	`

	w := httptest.NewRecorder()
	r := httptest.NewRequest("GET", "/test.asp", nil)

	processor := NewASPProcessor(&ASPProcessorConfig{
		RootDir:       "../www",
		ScriptTimeout: 30,
		DebugASP:      false,
	})

	err := processor.ExecuteASPFile(code, filepath.Join("..", "www", "test.asp"), w, r)
	if err != nil {
		t.Fatalf("ExecuteASPFile failed: %v", err)
	}

	output := w.Body.String()
	t.Logf("Output:\n%s", output)

	// Verify direct access works
	if !strings.Contains(output, "Direct access - name: field1") {
		t.Errorf("Direct dictionary access failed")
	}

	// Verify access via variable works
	if !strings.Contains(output, "Via variable - name: field1") {
		t.Errorf("Dictionary access via variable failed")
	}

	// Verify access from array works
	if !strings.Contains(output, "From array - name: field1") {
		t.Errorf("Dictionary access from array failed")
	}

	// Verify access in For Each loop works
	if !strings.Contains(output, "ForEach - name: field1") {
		t.Errorf("Dictionary access in For Each loop failed - this is the core issue")
	}
}
