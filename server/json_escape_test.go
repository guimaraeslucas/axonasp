package server

import (
	"net/http/httptest"
	"path/filepath"
	"strings"
	"testing"
)

// Test VBScript + operator behavior with strings
// In VBScript, when BOTH operands are strings, + performs concatenation
// This is critical for aspLite's JSON.escape function which uses + for string building
func TestPlusOperatorStringConcatenation(t *testing.T) {
	const page = `<%
' Test 1: Two string literals with +
Dim result1 : result1 = "Hello" + "World"
Response.Write "TEST1=" & result1 & "|"

' Test 2: String + Hex result (Hex returns string)
Dim result2 : result2 = "\u00" + Hex(47)
Response.Write "TEST2=" & result2 & "|"

' Test 3: String + Right result (Right returns string)
Dim result3 : result3 = "\u00" + Right("002F", 2)
Response.Write "TEST3=" & result3 & "|"

' Test 4: Number + Number (should add numerically)
Dim result4 : result4 = 10 + 20
Response.Write "TEST4=" & result4 & "|"

' Test 5: String that looks numeric + String that looks numeric - should concatenate
Dim result5 : result5 = "10" + "20"
Response.Write "TEST5=" & result5 & "|"
%>`

	req := httptest.NewRequest("GET", "http://example.com/test_plus.asp", nil)
	rec := httptest.NewRecorder()

	processor := NewASPProcessor(&ASPProcessorConfig{RootDir: filepath.Join("..", "www"), ScriptTimeout: 10})
	if err := processor.ExecuteASPFile(page, filepath.Join("..", "www", "test_plus.asp"), rec, req); err != nil {
		t.Fatalf("ExecuteASPFile returned error: %v", err)
	}

	body := rec.Body.String()
	t.Logf("body: %s", body)

	// Test 1: Basic string + string = concatenation
	if !strings.Contains(body, "TEST1=HelloWorld|") {
		t.Errorf("expected TEST1=HelloWorld|, got body: %s", body)
	}

	// Test 2: String + Hex result = concatenation "\u00" + "2F" = "\u002F"
	if !strings.Contains(body, "TEST2=\\u002F|") {
		t.Errorf("expected TEST2=\\u002F|, got body: %s", body)
	}

	// Test 3: String + Right result = concatenation
	if !strings.Contains(body, "TEST3=\\u002F|") {
		t.Errorf("expected TEST3=\\u002F|, got body: %s", body)
	}

	// Test 4: Number + Number = numeric addition
	if !strings.Contains(body, "TEST4=30|") {
		t.Errorf("expected TEST4=30|, got body: %s", body)
	}

	// Test 5: Two numeric strings + = concatenation (VBScript behavior)
	if !strings.Contains(body, "TEST5=1020|") {
		t.Errorf("expected TEST5=1020|, got body: %s", body)
	}
}

// Test that JSON escape function handles forward slashes correctly
// In JSON, forward slash (/) can optionally be escaped as \u002F or \/
// The aspLite JSON.escape function escapes it with \u002F
func TestJSONEscapeForwardSlash(t *testing.T) {
	const page = `<!-- #include file="asplite/asplite.asp"-->
<%
Response.Write "ASCW_SLASH=" & AscW("/") & "|"
Response.Write "HEX_47=" & Hex(47) & "|"
Response.Write "PADLEFT_2F=" & aspl.padLeft(Hex(47), 2, 0) & "|"
%>`

	req := httptest.NewRequest("GET", "http://example.com/test_json_escape.asp", nil)
	rec := httptest.NewRecorder()

	processor := NewASPProcessor(&ASPProcessorConfig{RootDir: filepath.Join("..", "www"), ScriptTimeout: 10})
	if err := processor.ExecuteASPFile(page, filepath.Join("..", "www", "test_json_escape.asp"), rec, req); err != nil {
		t.Fatalf("ExecuteASPFile returned error: %v", err)
	}

	body := rec.Body.String()
	t.Logf("body: %s", body)

	// AscW("/") should be 47
	if !strings.Contains(body, "ASCW_SLASH=47|") {
		t.Errorf("expected ASCW_SLASH=47, got body: %s", body)
	}

	// Hex(47) should be "2F"
	if !strings.Contains(body, "HEX_47=2F|") {
		t.Errorf("expected HEX_47=2F, got body: %s", body)
	}

	// padLeft("2F", 2, 0) should be "2F"
	if !strings.Contains(body, "PADLEFT_2F=2F|") {
		t.Errorf("expected PADLEFT_2F=2F, got body: %s", body)
	}
}

// Test the complete JSON escapesequence function with forward slash
func TestJSONEscapeSequence(t *testing.T) {
	const page = `<!-- #include file="asplite/asplite.asp"-->
<%
' Test what the json.escape function produces for "/"
Dim escaped : escaped = aspl.json.escape("/")
Response.Write "ESCAPED_SLASH=" & escaped & "|"

' Test the full string
Dim fullStr : fullStr = aspl.json.escape("plain/text")
Response.Write "ESCAPED_FULL=" & fullStr & "|"
%>`

	req := httptest.NewRequest("GET", "http://example.com/test_json_escape2.asp", nil)
	rec := httptest.NewRecorder()

	processor := NewASPProcessor(&ASPProcessorConfig{RootDir: filepath.Join("..", "www"), ScriptTimeout: 10})
	if err := processor.ExecuteASPFile(page, filepath.Join("..", "www", "test_json_escape2.asp"), rec, req); err != nil {
		t.Fatalf("ExecuteASPFile returned error: %v", err)
	}

	body := rec.Body.String()
	t.Logf("body: %s", body)

	// The escape function should produce "\u002F" for "/"
	if !strings.Contains(body, `ESCAPED_SLASH=\u002F|`) {
		t.Errorf("expected ESCAPED_SLASH=\\u002F|, got body: %s", body)
	}

	// For "plain/text" should produce "plain\u002Ftext"
	if !strings.Contains(body, `ESCAPED_FULL=plain\u002Ftext|`) {
		t.Errorf("expected ESCAPED_FULL=plain\\u002Ftext|, got body: %s", body)
	}
}

// Test the form.field("plain") with HTML content
func TestFormFieldPlainWithHTML(t *testing.T) {
	const page = `<!-- #include file="asplite/asplite.asp"-->
<%
dim form : set form=aspl.form

dim plain : set plain=form.field("plain")
plain.add "html", "This adds plain/text or <u>HTML</u>. Check the console (f12)!"

' Just test the dictionary value directly
Response.Write "DICT_HTML=" & plain("html") & "|"
%>`

	req := httptest.NewRequest("GET", "http://example.com/test_form_plain.asp?asplEvent=test", nil)
	rec := httptest.NewRecorder()

	processor := NewASPProcessor(&ASPProcessorConfig{RootDir: filepath.Join("..", "www"), ScriptTimeout: 10})
	if err := processor.ExecuteASPFile(page, filepath.Join("..", "www", "test_form_plain.asp"), rec, req); err != nil {
		t.Fatalf("ExecuteASPFile returned error: %v", err)
	}

	body := rec.Body.String()
	t.Logf("body: %s", body)

	// The dictionary should preserve the original string
	expected := `DICT_HTML=This adds plain/text or <u>HTML</u>. Check the console (f12)!|`
	if !strings.Contains(body, expected) {
		t.Errorf("expected %s, got body: %s", expected, body)
	}
}

// Test that the full form.build with plain.inc produces correct JSON
// This is the actual use case from the ebook/plain.inc file
func TestFormBuildPlainInc(t *testing.T) {
	const page = `<!-- #include file="asplite/asplite.asp"-->
<%
dim form : set form=aspl.form

dim plain : set plain=form.field("plain")
plain.add "html", "This adds plain/text or <u>HTML</u>. Check the console (f12)!"

dim script : set script=form.field("script")
script.add "text", "console.log('Add JavaScripts');"

form.build
%>`

	req := httptest.NewRequest("GET", "http://example.com/test_form_build.asp?asplEvent=plain", nil)
	rec := httptest.NewRecorder()

	processor := NewASPProcessor(&ASPProcessorConfig{RootDir: filepath.Join("..", "www"), ScriptTimeout: 10})
	if err := processor.ExecuteASPFile(page, filepath.Join("..", "www", "test_form_build.asp"), rec, req); err != nil {
		t.Fatalf("ExecuteASPFile returned error: %v", err)
	}

	body := rec.Body.String()
	t.Logf("body: %s", body)

	// The JSON should contain the properly escaped forward slash
	// "plain/text" should appear as "plain\/text" or "plain\u002Ftext" in JSON
	if strings.Contains(body, "plain0text") {
		t.Errorf("body incorrectly contains plain0text (0 instead of escaped slash): %s", body)
	}

	// Should contain the HTML content (escaped for JSON)
	if !strings.Contains(body, "plain") && !strings.Contains(body, "HTML") {
		t.Errorf("body missing expected content: %s", body)
	}
}
