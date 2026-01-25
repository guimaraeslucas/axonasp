package server

import (
	"net/http/httptest"
	"os"
	"path/filepath"
	"strings"
	"testing"
)

// Regression test: Request("key") should resolve QueryString values.
func TestRequestDefaultCollectionUsesQueryString(t *testing.T) {
	const aspPage = `<% Response.Write Request("asplEvent") %>`
	req := httptest.NewRequest("GET", "http://example.com/test.asp?asplEvent=fields", nil)
	rec := httptest.NewRecorder()

	processor := NewASPProcessor(&ASPProcessorConfig{RootDir: "./www", ScriptTimeout: 5})
	if err := processor.ExecuteASPFile(aspPage, "./www/test.asp", rec, req); err != nil {
		t.Fatalf("ExecuteASPFile returned error: %v", err)
	}

	body := rec.Body.String()
	if body != "fields" {
		t.Fatalf("expected Request default to return QueryString value, got %q", body)
	}
}

// End-to-end check: aspLite should honor asplEvent query parameter and return JSON instead of the full page.
func TestAspLiteAsplEventFieldsReturnsJSON(t *testing.T) {
	const inlinePage = `<%
dim asplEvent : asplEvent = Request("asplEvent")
if asplEvent <> "" then
    Response.ContentType = "application/json"
	Response.Write "{""asplEvent"":""" & asplEvent & """}"
else
    Response.Write "no event"
end if
%>`

	req := httptest.NewRequest("GET", "http://example.com/default.asp?asplEvent=fields", nil)
	rec := httptest.NewRecorder()

	processor := NewASPProcessor(&ASPProcessorConfig{RootDir: filepath.Join("..", "www"), ScriptTimeout: 10})
	if err := processor.ExecuteASPFile(inlinePage, filepath.Join("..", "www", "default.asp"), rec, req); err != nil {
		t.Fatalf("ExecuteASPFile returned error: %v", err)
	}

	contentType := rec.Header().Get("Content-Type")
	if !strings.Contains(strings.ToLower(contentType), "application/json") {
		t.Fatalf("expected JSON content type, got %q", contentType)
	}

	body := rec.Body.String()
	if body != `{"asplEvent":"fields"}` {
		t.Fatalf("expected JSON body with asplEvent, got: %s", body)
	}
}

// Ensures aspLite include execution happens when asplEvent is provided (smoke test on a small page).
func TestAspLiteDynamicPageRunsEventInclude(t *testing.T) {
	pagePath := filepath.Join("..", "www", "test_asplite_dynamic.asp")
	pageBytes, err := os.ReadFile(pagePath)
	if err != nil {
		t.Fatalf("failed to read test page: %v", err)
	}

	req := httptest.NewRequest("GET", "http://example.com/test_asplite_dynamic.asp?asplEvent=fields", nil)
	rec := httptest.NewRecorder()

	processor := NewASPProcessor(&ASPProcessorConfig{RootDir: filepath.Join("..", "www"), ScriptTimeout: 10})
	if err := processor.ExecuteASPFile(string(pageBytes), pagePath, rec, req); err != nil {
		t.Fatalf("ExecuteASPFile returned error: %v", err)
	}

	body := rec.Body.String()
	if !strings.Contains(body, "Found asplEvent") {
		t.Fatalf("expected event branch to run, got body: %s", body)
	}
}

// Full aspLite regression: test page should execute ebook fields and return JSON.
func TestAspLiteEbookFieldsReturnsJSON(t *testing.T) {
	pagePath := filepath.Join("..", "www", "test_asplite_ebook_fields.asp")
	pageBytes, err := os.ReadFile(pagePath)
	if err != nil {
		t.Fatalf("failed to read test_asplite_ebook_fields.asp: %v", err)
	}

	req := httptest.NewRequest("GET", "http://example.com/test_asplite_ebook_fields.asp?asplEvent=fields", nil)
	rec := httptest.NewRecorder()

	processor := NewASPProcessor(&ASPProcessorConfig{RootDir: filepath.Join("..", "www"), ScriptTimeout: 10})
	if err := processor.ExecuteASPFile(string(pageBytes), pagePath, rec, req); err != nil {
		t.Fatalf("ExecuteASPFile returned error: %v", err)
	}

	contentType := strings.ToLower(rec.Header().Get("Content-Type"))
	if !strings.Contains(contentType, "application/json") {
		t.Fatalf("expected JSON content type, got %q with body: %s", contentType, rec.Body.String())
	}

	body := rec.Body.String()
	if !strings.Contains(body, "\"aspForm\"") {
		t.Fatalf("expected aspForm JSON output, got body: %s", body)
	}
}

// Diagnose aspLite request helpers to ensure QueryString is visible through different access paths.
func TestAspLiteRequestHelpers(t *testing.T) {
	const page = `<!-- #include file="aspLite-master/aspLite/aspLite.asp"-->
<%
Response.Write "QS=" & Request.QueryString("asplEvent") & "\n"
Response.Write "REQ=" & Request("asplEvent") & "\n"
Response.Write "RTYPE=" & TypeName(Request) & "\n"
	Response.Write "FORMVAL=" & Request.Form("asplEvent") & "\n"
	Response.Write "QSVALUE=" & Request.QueryString("asplEvent") & "\n"
	Response.Write "ISFORM=" & aspl.isEmpty(Request.Form("asplEvent")) & "\n"
	Response.Write "ISQS=" & aspl.isEmpty(Request.QueryString("asplEvent")) & "\n"
Response.Write "REQ2=" & request("asplEvent") & "\n"
	dim tmp
	dim branch
	if not aspl.isEmpty(request.form("asplEvent")) then
		branch="form"
		tmp=request.form("asplEvent")
	elseif aspl.isEmpty(request.querystring("asplEvent")) then
		branch="qs_empty"
		tmp=request.querystring("asplEvent")
	else
		branch="else"
		tmp=request("asplEvent")
	end if
	Response.Write "BRANCH=" & branch & "\n"
	Response.Write "MANUAL=" & tmp & "\n"
	Response.Write "ERRNUM=" & Err.Number & "\n"
	Response.Write "GETREQ=" & aspl.getRequest("asplEvent") & "\n"
%>`

	req := httptest.NewRequest("GET", "http://example.com/test_req.asp?asplEvent=fields", nil)
	rec := httptest.NewRecorder()

	processor := NewASPProcessor(&ASPProcessorConfig{RootDir: filepath.Join("..", "www"), ScriptTimeout: 10})
	if err := processor.ExecuteASPFile(page, filepath.Join("..", "www", "test_req.asp"), rec, req); err != nil {
		t.Fatalf("ExecuteASPFile returned error: %v", err)
	}

	body := rec.Body.String()
	checks := []string{
		"QS=fields",
		"REQ=fields",
		"REQ2=fields",
		"QSVALUE=fields",
		"GETREQ=fields",
	}
	for _, expect := range checks {
		if !strings.Contains(body, expect) {
			t.Fatalf("missing %q in output. Full body: %s", expect, body)
		}
	}
}

// Sanity check: aspLite global instance should be available via case-insensitive name.
func TestAspLiteInstanceCaseInsensitiveAccess(t *testing.T) {
	const page = `<!-- #include file="aspLite-master/aspLite/aspLite.asp"-->
<%
Response.Write "ASPL_TYPE=" & TypeName(aspl) & "\n"
Response.Write "ASPL_UPPER=" & TypeName(aspL) & "\n"
%>`

	req := httptest.NewRequest("GET", "http://example.com/test_aspl_instance.asp", nil)
	rec := httptest.NewRecorder()

	processor := NewASPProcessor(&ASPProcessorConfig{RootDir: filepath.Join("..", "www"), ScriptTimeout: 10})
	if err := processor.ExecuteASPFile(page, filepath.Join("..", "www", "test_aspl_instance.asp"), rec, req); err != nil {
		t.Fatalf("ExecuteASPFile returned error: %v", err)
	}

	body := rec.Body.String()
	if !strings.Contains(body, "ASPL_TYPE=Object") {
		t.Fatalf("expected aspl to be an object, got body: %s", body)
	}
	if !strings.Contains(body, "ASPL_UPPER=Object") {
		t.Fatalf("expected aspL to be an object, got body: %s", body)
	}
}

// Ensure aspLite default exec can execute .inc content via ExecuteGlobal.
func TestAspLiteDefaultExecRunsInclude(t *testing.T) {
	const page = `<!-- #include file="aspLite-master/aspLite/aspLite.asp"-->
<%
aspl("test_aspl_exec.inc")
%>`

	req := httptest.NewRequest("GET", "http://example.com/test_aspl_exec.asp", nil)
	rec := httptest.NewRecorder()

	processor := NewASPProcessor(&ASPProcessorConfig{RootDir: filepath.Join("..", "www"), ScriptTimeout: 10})
	if err := processor.ExecuteASPFile(page, filepath.Join("..", "www", "test_aspl_exec.asp"), rec, req); err != nil {
		t.Fatalf("ExecuteASPFile returned error: %v", err)
	}

	body := rec.Body.String()
	if !strings.Contains(body, "ASPL_EXEC_OK") {
		t.Fatalf("expected include output, got body: %s", body)
	}
}

// Explicit call to aspLite exec should run include even if default dispatch fails.
func TestAspLiteExplicitExecRunsInclude(t *testing.T) {
	const page = `<!-- #include file="aspLite-master/aspLite/aspLite.asp"-->
<%
Call aspl.exec("test_aspl_exec.inc")
%>`

	req := httptest.NewRequest("GET", "http://example.com/test_aspl_exec_explicit.asp", nil)
	rec := httptest.NewRecorder()

	processor := NewASPProcessor(&ASPProcessorConfig{RootDir: filepath.Join("..", "www"), ScriptTimeout: 10})
	if err := processor.ExecuteASPFile(page, filepath.Join("..", "www", "test_aspl_exec_explicit.asp"), rec, req); err != nil {
		t.Fatalf("ExecuteASPFile returned error: %v", err)
	}

	body := rec.Body.String()
	if !strings.Contains(body, "ASPL_EXEC_OK") {
		t.Fatalf("expected explicit exec output, got body: %s", body)
	}
}

// ExecuteGlobal should run VBScript snippets in the current context.
func TestExecuteGlobalRunsInline(t *testing.T) {
	const page = `<% ExecuteGlobal "Response.Write ""EXEC_GLOBAL_OK""" %>`

	req := httptest.NewRequest("GET", "http://example.com/test_exec_global.asp", nil)
	rec := httptest.NewRecorder()

	processor := NewASPProcessor(&ASPProcessorConfig{RootDir: filepath.Join("..", "www"), ScriptTimeout: 10})
	if err := processor.ExecuteASPFile(page, filepath.Join("..", "www", "test_exec_global.asp"), rec, req); err != nil {
		t.Fatalf("ExecuteASPFile returned error: %v", err)
	}

	body := rec.Body.String()
	if !strings.Contains(body, "EXEC_GLOBAL_OK") {
		t.Fatalf("expected ExecuteGlobal output, got body: %s", body)
	}
}

// aspLite stream helper should read file contents via ADODB.Stream.
func TestAspLiteStreamLoadText(t *testing.T) {
	const page = `<!-- #include file="aspLite-master/aspLite/aspLite.asp"-->
<%
Response.Write aspl.loadText("test_aspl_exec.inc")
%>`

	req := httptest.NewRequest("GET", "http://example.com/test_aspl_loadtext.asp", nil)
	rec := httptest.NewRecorder()

	processor := NewASPProcessor(&ASPProcessorConfig{RootDir: filepath.Join("..", "www"), ScriptTimeout: 10})
	if err := processor.ExecuteASPFile(page, filepath.Join("..", "www", "test_aspl_loadtext.asp"), rec, req); err != nil {
		t.Fatalf("ExecuteASPFile returned error: %v", err)
	}

	body := rec.Body.String()
	if !strings.Contains(body, "ASPL_EXEC_OK") {
		t.Fatalf("expected loadText to include file contents, got body: %s", body)
	}
}

// Ensure regular method dispatch on aspLite instance works.
func TestAspLiteMethodDispatch(t *testing.T) {
	const page = `<!-- #include file="aspLite-master/aspLite/aspLite.asp"-->
<%
Response.Write "CONVERT=" & aspl.convertStr("x") & "|EMPTY=" & aspl.isEmpty("") & "|TRIM=" & Trim("") & "|LEN=" & Len("")
%>`

	req := httptest.NewRequest("GET", "http://example.com/test_aspl_method.asp", nil)
	rec := httptest.NewRecorder()

	processor := NewASPProcessor(&ASPProcessorConfig{RootDir: filepath.Join("..", "www"), ScriptTimeout: 10})
	if err := processor.ExecuteASPFile(page, filepath.Join("..", "www", "test_aspl_method.asp"), rec, req); err != nil {
		t.Fatalf("ExecuteASPFile returned error: %v", err)
	}

	body := rec.Body.String()
	if !strings.Contains(body, "CONVERT=x") {
		t.Fatalf("expected convertStr output, got body: %s", body)
	}
	if !strings.Contains(body, "EMPTY=True") {
		t.Fatalf("expected isEmpty(True) result, got body: %s", body)
	}
}

// ExecuteGlobal should work with aspLite loadText + removeCRB output.
func TestExecuteGlobalFromLoadText(t *testing.T) {
	const page = `<!-- #include file="aspLite-master/aspLite/aspLite.asp"-->
<%
Dim code
code = aspl.removeCRB(aspl.loadText("test_aspl_exec.inc"))
ExecuteGlobal code
%>`

	req := httptest.NewRequest("GET", "http://example.com/test_exec_global_loadtext.asp", nil)
	rec := httptest.NewRecorder()

	processor := NewASPProcessor(&ASPProcessorConfig{RootDir: filepath.Join("..", "www"), ScriptTimeout: 10})
	if err := processor.ExecuteASPFile(page, filepath.Join("..", "www", "test_exec_global_loadtext.asp"), rec, req); err != nil {
		t.Fatalf("ExecuteASPFile returned error: %v", err)
	}

	body := rec.Body.String()
	if !strings.Contains(body, "ASPL_EXEC_OK") {
		t.Fatalf("expected ExecuteGlobal output from loadText, got body: %s", body)
	}
}

// Ensure aspLite removeCRB strips ASP tags as expected.
func TestAspLiteRemoveCRB(t *testing.T) {
	const page = `<!-- #include file="aspLite-master/aspLite/aspLite.asp"-->
<%
Response.Write aspl.removeCRB(aspl.loadText("test_aspl_exec.inc"))
%>`

	req := httptest.NewRequest("GET", "http://example.com/test_aspl_removecrb.asp", nil)
	rec := httptest.NewRecorder()

	processor := NewASPProcessor(&ASPProcessorConfig{RootDir: filepath.Join("..", "www"), ScriptTimeout: 10})
	if err := processor.ExecuteASPFile(page, filepath.Join("..", "www", "test_aspl_removecrb.asp"), rec, req); err != nil {
		t.Fatalf("ExecuteASPFile returned error: %v", err)
	}

	body := rec.Body.String()
	if strings.Contains(body, "<%") || strings.Contains(body, "%>") {
		t.Fatalf("expected removeCRB to strip ASP tags, got body: %s", body)
	}
	if !strings.Contains(body, "ASPL_EXEC_OK") {
		t.Fatalf("expected removeCRB to preserve code content, got body: %s", body)
	}
}

// Ensure aspLite class methods are not leaking to global scope.
func TestAspLiteLoadTextNotGlobal(t *testing.T) {
	const page = `<!-- #include file="aspLite-master/aspLite/aspLite.asp"-->
<%
On Error Resume Next
Response.Write loadText("test_aspl_exec.inc")
Response.Write "|ERR=" & Err.Number
%>`

	req := httptest.NewRequest("GET", "http://example.com/test_aspl_loadtext_global.asp", nil)
	rec := httptest.NewRecorder()

	processor := NewASPProcessor(&ASPProcessorConfig{RootDir: filepath.Join("..", "www"), ScriptTimeout: 10})
	if err := processor.ExecuteASPFile(page, filepath.Join("..", "www", "test_aspl_loadtext_global.asp"), rec, req); err != nil {
		t.Fatalf("ExecuteASPFile returned error: %v", err)
	}

	body := rec.Body.String()
	if strings.Contains(body, "ASPL_EXEC_OK") {
		t.Fatalf("expected loadText to be scoped to aspLite instance, got body: %s", body)
	}
	if !strings.Contains(body, "ERR=") {
		t.Fatalf("expected error output when calling global loadText, got body: %s", body)
	}
}

// Sanity check: class Sub methods should be callable on instances.
func TestClassSubMethodCall(t *testing.T) {
	const page = `<%
Class TestClass
	Public Sub Exec(path)
		Response.Write "SUB_OK"
	End Sub
End Class

Dim t
Set t = New TestClass
Call t.Exec("x")
%>`

	req := httptest.NewRequest("GET", "http://example.com/test_class_sub.asp", nil)
	rec := httptest.NewRecorder()

	processor := NewASPProcessor(&ASPProcessorConfig{RootDir: filepath.Join("..", "www"), ScriptTimeout: 10})
	if err := processor.ExecuteASPFile(page, filepath.Join("..", "www", "test_class_sub.asp"), rec, req); err != nil {
		t.Fatalf("ExecuteASPFile returned error: %v", err)
	}

	body := rec.Body.String()
	if !strings.Contains(body, "SUB_OK") {
		t.Fatalf("expected class Sub output, got body: %s", body)
	}
}

// Default Sub should be callable explicitly and via default dispatch.
func TestClassDefaultSubDispatch(t *testing.T) {
	const page = `<%
Class DefaultClass
	Public Default Sub Exec(path)
		Response.Write "DEFAULT_SUB_OK"
	End Sub
End Class

Dim t
Set t = New DefaultClass
Call t.Exec("x")
t "y"
%>`

	req := httptest.NewRequest("GET", "http://example.com/test_default_sub.asp", nil)
	rec := httptest.NewRecorder()

	processor := NewASPProcessor(&ASPProcessorConfig{RootDir: filepath.Join("..", "www"), ScriptTimeout: 10})
	if err := processor.ExecuteASPFile(page, filepath.Join("..", "www", "test_default_sub.asp"), rec, req); err != nil {
		t.Fatalf("ExecuteASPFile returned error: %v", err)
	}

	body := rec.Body.String()
	if strings.Count(body, "DEFAULT_SUB_OK") != 2 {
		t.Fatalf("expected default sub to run twice, got body: %s", body)
	}
}

// Class function named isEmpty should not be confused with built-in IsEmpty.
func TestClassIsEmptyMethod(t *testing.T) {
	const page = `<%
Class NamedMethods
	Public Function isEmpty(val)
		If val = "" Then
			isEmpty = True
		Else
			isEmpty = False
		End If
	End Function
End Class

Dim o
Set o = New NamedMethods
Response.Write "EMPTY=" & o.isEmpty("")
%>`

	req := httptest.NewRequest("GET", "http://example.com/test_class_isempty.asp", nil)
	rec := httptest.NewRecorder()

	processor := NewASPProcessor(&ASPProcessorConfig{RootDir: filepath.Join("..", "www"), ScriptTimeout: 10})
	if err := processor.ExecuteASPFile(page, filepath.Join("..", "www", "test_class_isempty.asp"), rec, req); err != nil {
		t.Fatalf("ExecuteASPFile returned error: %v", err)
	}

	body := rec.Body.String()
	if !strings.Contains(body, "EMPTY=True") {
		t.Fatalf("expected class isEmpty method to return True, got body: %s", body)
	}
}

// Reproduce aspLite isEmpty logic in a standalone class for comparison.
func TestClassAspLiteIsEmptyLogic(t *testing.T) {
	const page = `<%
Class AspLiteLike
	Public Function isEmpty(ByVal value)
		On Error Resume Next
		isEmpty = False
		If IsEmpty(value) Then
			isEmpty = True
		ElseIf IsNull(value) Then
			isEmpty = True
		Else
			If Trim(value) = "" Then isEmpty = True
		End If
		On Error GoTo 0
	End Function
End Class

Dim o
Set o = New AspLiteLike
Response.Write "EMPTY=" & o.isEmpty("")
%>`

	req := httptest.NewRequest("GET", "http://example.com/test_class_isempty_logic.asp", nil)
	rec := httptest.NewRecorder()

	processor := NewASPProcessor(&ASPProcessorConfig{RootDir: filepath.Join("..", "www"), ScriptTimeout: 10})
	if err := processor.ExecuteASPFile(page, filepath.Join("..", "www", "test_class_isempty_logic.asp"), rec, req); err != nil {
		t.Fatalf("ExecuteASPFile returned error: %v", err)
	}

	body := rec.Body.String()
	if !strings.Contains(body, "EMPTY=True") {
		t.Fatalf("expected aspLite isEmpty logic to return True, got body: %s", body)
	}
}

// Basic Trim + equality check should work in VBScript expressions.
func TestTrimEquality(t *testing.T) {
	const page = `<%
If Trim("") = "" Then
	Response.Write "TRIM_EQ_OK"
Else
	Response.Write "TRIM_EQ_FAIL"
End If
%>`

	req := httptest.NewRequest("GET", "http://example.com/test_trim_eq.asp", nil)
	rec := httptest.NewRecorder()

	processor := NewASPProcessor(&ASPProcessorConfig{RootDir: filepath.Join("..", "www"), ScriptTimeout: 10})
	if err := processor.ExecuteASPFile(page, filepath.Join("..", "www", "test_trim_eq.asp"), rec, req); err != nil {
		t.Fatalf("ExecuteASPFile returned error: %v", err)
	}

	body := rec.Body.String()
	if !strings.Contains(body, "TRIM_EQ_OK") {
		t.Fatalf("expected Trim equality to pass, got body: %s", body)
	}
}

// Class method defined with bracket notation [isEmpty] should work like regular methods
func TestClassBracketMethodName(t *testing.T) {
	const page = `<%
Class TestBracket
	Public Function [isEmpty](ByVal value)
		[isEmpty] = False
		If Trim(value) = "" Then
			[isEmpty] = True
		End If
	End Function
End Class

Dim o
Set o = New TestBracket
Response.Write "BRACKET=" & o.isEmpty("")
%>`

	req := httptest.NewRequest("GET", "http://example.com/test_bracket.asp", nil)
	rec := httptest.NewRecorder()

	processor := NewASPProcessor(&ASPProcessorConfig{RootDir: filepath.Join("..", "www"), ScriptTimeout: 10})
	if err := processor.ExecuteASPFile(page, filepath.Join("..", "www", "test_bracket.asp"), rec, req); err != nil {
		t.Fatalf("ExecuteASPFile returned error: %v", err)
	}

	body := rec.Body.String()
	if !strings.Contains(body, "BRACKET=True") {
		t.Fatalf("expected bracket method to return True, got body: %s", body)
	}
}

// Test bracket method in aspLite-like context with global instance  
func TestClassBracketMethodGlobalInstance(t *testing.T) {
	const page = `<%
Class cls_test
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

Dim myObj
Set myObj = New cls_test
Response.Write "GLOBAL_BRACKET=" & myObj.isEmpty("")
%>`

	req := httptest.NewRequest("GET", "http://example.com/test_global_bracket.asp", nil)
	rec := httptest.NewRecorder()

	processor := NewASPProcessor(&ASPProcessorConfig{RootDir: filepath.Join("..", "www"), ScriptTimeout: 10})
	if err := processor.ExecuteASPFile(page, filepath.Join("..", "www", "test_global_bracket.asp"), rec, req); err != nil {
		t.Fatalf("ExecuteASPFile returned error: %v", err)
	}

	body := rec.Body.String()
	t.Logf("body: %s", body)
	if !strings.Contains(body, "GLOBAL_BRACKET=True") {
		t.Fatalf("expected bracket method on global instance to return True, got body: %s", body)
	}
}

// Test ElseIf in a simple context
func TestElseIfBasic(t *testing.T) {
	const page = `<%
Dim x
x = 3
If x = 1 Then
	Response.Write "ONE"
ElseIf x = 2 Then
	Response.Write "TWO"
ElseIf x = 3 Then
	Response.Write "THREE"
Else
	Response.Write "OTHER"
End If
%>`

	req := httptest.NewRequest("GET", "http://example.com/test_elseif.asp", nil)
	rec := httptest.NewRecorder()

	processor := NewASPProcessor(&ASPProcessorConfig{RootDir: filepath.Join("..", "www"), ScriptTimeout: 10})
	if err := processor.ExecuteASPFile(page, filepath.Join("..", "www", "test_elseif.asp"), rec, req); err != nil {
		t.Fatalf("ExecuteASPFile returned error: %v", err)
	}

	body := rec.Body.String()
	t.Logf("body: %s", body)
	if !strings.Contains(body, "THREE") {
		t.Fatalf("expected THREE, got body: %s", body)
	}
}

// aspLite executeASP should run inline ASP code passed as string.
func TestAspLiteExecuteASP(t *testing.T) {
	const page = `<!-- #include file="aspLite-master/aspLite/aspLite.asp"-->
<%
Call aspl.executeASP("<% Response.Write ""EXEC_ASP_OK"" %>")
%>`

	req := httptest.NewRequest("GET", "http://example.com/test_aspl_executeasp.asp", nil)
	rec := httptest.NewRecorder()

	processor := NewASPProcessor(&ASPProcessorConfig{RootDir: filepath.Join("..", "www"), ScriptTimeout: 10})
	if err := processor.ExecuteASPFile(page, filepath.Join("..", "www", "test_aspl_executeasp.asp"), rec, req); err != nil {
		t.Fatalf("ExecuteASPFile returned error: %v", err)
	}

	body := rec.Body.String()
	if !strings.Contains(body, "EXEC_ASP_OK") {
		t.Fatalf("expected executeASP output, got body: %s", body)
	}
}

// Test aspl default sub exec() is triggered when calling aspl("path")
func TestAsplDefaultExec(t *testing.T) {
	// Test aspLite exec method via explicit call
	const asplExecExplicit = `<!-- #include file="aspLite-master/aspLite/aspLite.asp"-->
<%
Response.Write "START|"
Call aspl.exec("test_aspl_exec.inc")
Response.Write "|END"
%>`

	req := httptest.NewRequest("GET", "http://example.com/test.asp", nil)
	rec := httptest.NewRecorder()

	processor := NewASPProcessor(&ASPProcessorConfig{RootDir: filepath.Join("..", "www"), ScriptTimeout: 10})
	if err := processor.ExecuteASPFile(asplExecExplicit, filepath.Join("..", "www", "test.asp"), rec, req); err != nil {
		t.Fatalf("ExecuteASPFile returned error: %v", err)
	}

	body := rec.Body.String()
	t.Logf("Explicit exec body: %s", body)
	
	if !strings.Contains(body, "ASPL_EXEC_OK") {
		t.Fatalf("expected aspl.exec() explicit output, got body: %s", body)
	}

	// Test aspLite default sub via aspl("path") syntax
	const asplExecDefault = `<!-- #include file="aspLite-master/aspLite/aspLite.asp"-->
<%
Response.Write "START|"
aspl "test_aspl_exec.inc"
Response.Write "|END"
%>`

	req2 := httptest.NewRequest("GET", "http://example.com/test.asp", nil)
	rec2 := httptest.NewRecorder()

	processor2 := NewASPProcessor(&ASPProcessorConfig{RootDir: filepath.Join("..", "www"), ScriptTimeout: 10})
	if err := processor2.ExecuteASPFile(asplExecDefault, filepath.Join("..", "www", "test.asp"), rec2, req2); err != nil {
		t.Fatalf("ExecuteASPFile returned error: %v", err)
	}

	body2 := rec2.Body.String()
	t.Logf("Default exec body: %s", body2)
	
	if !strings.Contains(body2, "ASPL_EXEC_OK") {
		t.Fatalf("expected aspl(...) default sub output, got body: %s", body2)
	}
}
