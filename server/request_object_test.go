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
