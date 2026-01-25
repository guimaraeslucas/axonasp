package server

import (
	"net/http/httptest"
	"path/filepath"
	"strings"
	"testing"
)

// Test that bracketed function calls inside class methods work correctly
func TestBracketedFunctionInClass(t *testing.T) {
	const page = `<%
Class TestClass
	' Define a method with same name as built-in function
	Public Function [isEmpty](v)
		[isEmpty] = "CUSTOM:" & v
	End Function
	
	Public Function CallIt(v)
		' This should call our class method, not VBScript IsEmpty
		CallIt = [isEmpty](v)
	End Function
End Class

Dim obj : Set obj = New TestClass
Dim result : result = obj.CallIt("test")
Response.Write "RESULT=" & result & "|"
%>`

	req := httptest.NewRequest("GET", "http://example.com/test.asp", nil)
	rec := httptest.NewRecorder()

	processor := NewASPProcessor(&ASPProcessorConfig{RootDir: filepath.Join("..", "www"), ScriptTimeout: 10})
	if err := processor.ExecuteASPFile(page, filepath.Join("..", "www", "test.asp"), rec, req); err != nil {
		t.Fatalf("ExecuteASPFile returned error: %v", err)
	}

	body := rec.Body.String()
	t.Logf("body: %s", body)

	// Should call the class method, returning "CUSTOM:test"
	if !strings.Contains(body, "RESULT=CUSTOM:test|") {
		t.Fatalf("expected RESULT=CUSTOM:test|, got body: %s", body)
	}
}

// Test the ebook.asp flow: aspl.getRequest() and aspl.isEmpty()
func TestEbookAsplEventFlow(t *testing.T) {
	// Test the EXACT getRequest logic with bracket isEmpty notation
	const page = `
<%
Class cls_test_request
	Private multipart
	
	Private Sub Class_Initialize
		multipart = false
	End Sub
	
	Public Function [isEmpty](v)
		Response.Write "  isEmpty called with type=" & TypeName(v) & " value=[" & v & "]|"
		If IsNull(v) Then
			Response.Write "  isEmpty returning True (isNull)|"
			[isEmpty] = True
		ElseIf v = "" Then
			Response.Write "  isEmpty returning True (empty string)|"
			[isEmpty] = True
		Else
			Response.Write "  isEmpty returning False|"
			[isEmpty] = False
		End If
	End Function
	
	Public Function getRequest(value)
		Response.Write "ENTERING getRequest|"
		
		dim formValue : formValue = request.form(value)
		dim qsValue : qsValue = request.querystring(value)
		Response.Write "FORM_VALUE=[" & formValue & "]|"
		Response.Write "QS_VALUE=[" & qsValue & "]|"
		Response.Write "FORM_TYPENAME=" & TypeName(formValue) & "|"
		
		dim formEmpty : formEmpty = [isEmpty](formValue)
		Response.Write "ISEMPTY_FORM=" & formEmpty & "|"
		
		if not formEmpty then
			Response.Write "BRANCH=form_not_empty|"
			getRequest=formValue
		elseif [isEmpty](qsValue) then
			Response.Write "BRANCH=qs_empty|"
			getRequest=qsValue
		else
			Response.Write "BRANCH=else|"
			getRequest=request(value)
		end if
		Response.Write "RESULT=" & getRequest & "|"
	End Function
End Class

Dim testAspl : Set testAspl = New cls_test_request

Response.Write "DEBUG START|"
dim asplEvent : asplEvent = testAspl.getRequest("asplEvent")
Response.Write "EVENT=" & asplEvent & "|"
Response.Write "DEBUG END"
%>`

	req := httptest.NewRequest("GET", "http://example.com/test.asp?asplEvent=fields", nil)
	rec := httptest.NewRecorder()

	processor := NewASPProcessor(&ASPProcessorConfig{RootDir: filepath.Join("..", "www"), ScriptTimeout: 10})
	if err := processor.ExecuteASPFile(page, filepath.Join("..", "www", "test.asp"), rec, req); err != nil {
		t.Fatalf("ExecuteASPFile returned error: %v", err)
	}

	body := rec.Body.String()
	t.Logf("body: %s", body)

	// Check that we got the event value through the class
	if !strings.Contains(body, "EVENT=fields|") {
		t.Fatalf("expected EVENT=fields, got body: %s", body)
	}
}

// Test with the actual aspLite include file
func TestWithRealAspLiteInclude(t *testing.T) {
	const page = `<!-- #include file="asplite/asplite.asp"-->
<%
Response.Write "DEBUG START|"
Response.Write "DIRECT_QS=" & Request.QueryString("asplEvent") & "|"
dim asplEvent : asplEvent=aspl.getRequest("asplEvent")
Response.Write "EVENT=" & asplEvent & "|"
dim isEmpty : isEmpty=aspl.isEmpty(asplEvent)
Response.Write "ISEMPTY=" & isEmpty & "|"
if not aspl.isEmpty(asplEvent) then
	Response.Write "BRANCH_TAKEN|"
else
	Response.Write "BRANCH_NOT_TAKEN|"
end if
Response.Write "DEBUG END"
%>`

	req := httptest.NewRequest("GET", "http://example.com/ebook.asp?asplEvent=fields", nil)
	rec := httptest.NewRecorder()

	processor := NewASPProcessor(&ASPProcessorConfig{RootDir: filepath.Join("..", "www"), ScriptTimeout: 10})
	if err := processor.ExecuteASPFile(page, filepath.Join("..", "www", "ebook.asp"), rec, req); err != nil {
		t.Fatalf("ExecuteASPFile returned error: %v", err)
	}

	body := rec.Body.String()
	t.Logf("body: %s", body)

	// Check that we got the event value
	if !strings.Contains(body, "EVENT=fields|") {
		t.Fatalf("expected EVENT=fields, got body: %s", body)
	}

	// Check that isEmpty returned False
	if !strings.Contains(body, "ISEMPTY=False|") {
		t.Fatalf("expected ISEMPTY=False, got body: %s", body)
	}

	// Check that the branch was taken
	if !strings.Contains(body, "BRANCH_TAKEN|") {
		t.Fatalf("expected BRANCH_TAKEN, got body: %s", body)
	}
}

// Test that the isEmpty method in aspLite class works correctly
func TestAsplIsEmptyMethod(t *testing.T) {
	// Simplified test of aspLite isEmpty method
	const page = `<%
Class TestAsplLike
	Public Function [isEmpty](ByVal value)
		On Error Resume Next
		isEmpty = False
		If IsNull(value) Then
			isEmpty = True
		Else
			If Trim(value) = "" Then isEmpty = True
		End If
		On Error GoTo 0
	End Function

	Public Function getRequest(key)
		getRequest = Request(key)
	End Function
End Class

Dim aspl : Set aspl = New TestAsplLike
Dim evt : evt = aspl.getRequest("asplEvent")
Response.Write "EVENT=" & evt & "|"
Response.Write "ISEMPTY=" & aspl.isEmpty(evt) & "|"
If Not aspl.isEmpty(evt) Then
	Response.Write "BRANCH_TAKEN"
Else
	Response.Write "BRANCH_NOT_TAKEN"
End If
%>`

	req := httptest.NewRequest("GET", "http://example.com/test.asp?asplEvent=fields", nil)
	rec := httptest.NewRecorder()

	processor := NewASPProcessor(&ASPProcessorConfig{RootDir: filepath.Join("..", "www"), ScriptTimeout: 10})
	if err := processor.ExecuteASPFile(page, filepath.Join("..", "www", "test.asp"), rec, req); err != nil {
		t.Fatalf("ExecuteASPFile returned error: %v", err)
	}

	body := rec.Body.String()
	t.Logf("body: %s", body)

	// Check that we got the event value
	if !strings.Contains(body, "EVENT=fields|") {
		t.Fatalf("expected EVENT=fields, got body: %s", body)
	}

	// Check that isEmpty returned False for "fields"
	if !strings.Contains(body, "ISEMPTY=False|") {
		t.Fatalf("expected ISEMPTY=False for non-empty string, got body: %s", body)
	}

	// Check that the branch was taken
	if !strings.Contains(body, "BRANCH_TAKEN") {
		t.Fatalf("expected BRANCH_TAKEN, got body: %s", body)
	}
}

// Test that bracket method names work with built-in function calls inside
func TestBracketMethodWithBuiltInIsEmpty(t *testing.T) {
	// Test that [isEmpty] method calling IsEmpty inside works correctly
	const page = `<%
Class TestClass
	Public Function [isEmpty](ByVal value)
		[isEmpty] = False
		If IsEmpty(value) Then
			[isEmpty] = True
		ElseIf IsNull(value) Then
			[isEmpty] = True
		Else
			If Trim(value) = "" Then [isEmpty] = True
		End If
	End Function
End Class

Dim obj : Set obj = New TestClass
Response.Write "EMPTY_STR=" & obj.isEmpty("") & "|"
Response.Write "NON_EMPTY=" & obj.isEmpty("hello") & "|"
%>`

	req := httptest.NewRequest("GET", "http://example.com/test.asp", nil)
	rec := httptest.NewRecorder()

	processor := NewASPProcessor(&ASPProcessorConfig{RootDir: filepath.Join("..", "www"), ScriptTimeout: 10})
	if err := processor.ExecuteASPFile(page, filepath.Join("..", "www", "test.asp"), rec, req); err != nil {
		t.Fatalf("ExecuteASPFile returned error: %v", err)
	}

	body := rec.Body.String()
	t.Logf("body: %s", body)

	// Empty string should return True
	if !strings.Contains(body, "EMPTY_STR=True|") {
		t.Fatalf("expected EMPTY_STR=True for empty string, got body: %s", body)
	}

	// Non-empty string should return False
	if !strings.Contains(body, "NON_EMPTY=False|") {
		t.Fatalf("expected NON_EMPTY=False for non-empty string, got body: %s", body)
	}
}

// Test IsNull and IsEmpty separately
func TestIsNullVsIsEmpty(t *testing.T) {
	const page = `<%
Response.Write "ISNULL_NULL=" & IsNull(Null) & "|"
Response.Write "ISEMPTY_NULL=" & IsEmpty(Null) & "|"
Response.Write "ISNULL_EMPTY=" & IsNull("") & "|"
Response.Write "ISEMPTY_EMPTY=" & IsEmpty("") & "|"
%>`

	req := httptest.NewRequest("GET", "http://example.com/test.asp", nil)
	rec := httptest.NewRecorder()

	processor := NewASPProcessor(&ASPProcessorConfig{RootDir: filepath.Join("..", "www"), ScriptTimeout: 10})
	if err := processor.ExecuteASPFile(page, filepath.Join("..", "www", "test.asp"), rec, req); err != nil {
		t.Fatalf("ExecuteASPFile returned error: %v", err)
	}

	body := rec.Body.String()
	t.Logf("body: %s", body)

	// IsNull(Null) should be True
	if !strings.Contains(body, "ISNULL_NULL=True|") {
		t.Fatalf("expected ISNULL_NULL=True, got body: %s", body)
	}

	// IsEmpty(Null) should be False (Null is not Empty)
	if !strings.Contains(body, "ISEMPTY_NULL=False|") {
		t.Fatalf("expected ISEMPTY_NULL=False, got body: %s", body)
	}

	// IsNull("") should be False
	if !strings.Contains(body, "ISNULL_EMPTY=False|") {
		t.Fatalf("expected ISNULL_EMPTY=False, got body: %s", body)
	}

	// IsEmpty("") should be False (empty string is not Empty in VBScript)
	if !strings.Contains(body, "ISEMPTY_EMPTY=False|") {
		t.Fatalf("expected ISEMPTY_EMPTY=False, got body: %s", body)
	}
}

// Test ElseIf inside a class method
func TestClassMethodElseIf(t *testing.T) {
	const page = `<%
Class TestClass
	Public Function Check(value)
		Check = "UNKNOWN"
		If IsEmpty(value) Then
			Check = "EMPTY"
		ElseIf IsNull(value) Then
			Check = "NULL"
		ElseIf value = "" Then
			Check = "BLANK"
		Else
			Check = "HAS_VALUE"
		End If
	End Function
End Class

Dim obj : Set obj = New TestClass
Response.Write "1=" & obj.Check("hello") & "|"
Response.Write "2=" & obj.Check("") & "|"
Response.Write "3=" & obj.Check(Null) & "|"
%>`

	req := httptest.NewRequest("GET", "http://example.com/test.asp", nil)
	rec := httptest.NewRecorder()

	processor := NewASPProcessor(&ASPProcessorConfig{RootDir: filepath.Join("..", "www"), ScriptTimeout: 10})
	if err := processor.ExecuteASPFile(page, filepath.Join("..", "www", "test.asp"), rec, req); err != nil {
		t.Fatalf("ExecuteASPFile returned error: %v", err)
	}

	body := rec.Body.String()
	t.Logf("body: %s", body)

	if !strings.Contains(body, "1=HAS_VALUE|") {
		t.Fatalf("expected 1=HAS_VALUE, got body: %s", body)
	}

	if !strings.Contains(body, "2=BLANK|") {
		t.Fatalf("expected 2=BLANK, got body: %s", body)
	}

	if !strings.Contains(body, "3=NULL|") {
		t.Fatalf("expected 3=NULL, got body: %s", body)
	}
}
