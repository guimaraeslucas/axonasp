package main

import (
	"os"
	"strings"
	"testing"
)

func TestInjectAppScript_BeforeBodyClose(t *testing.T) {
	html := []byte(`<html><body><h1>Hello</h1></body></html>`)
	result := injectAppScript(html)
	s := string(result)
	if !strings.Contains(s, "__heartbeat__") {
		t.Error("expected injected script to contain heartbeat fetch")
	}
	// Script should appear before </body>
	bodyIdx := strings.Index(strings.ToLower(s), "</body>")
	scriptIdx := strings.Index(s, "__heartbeat__")
	if scriptIdx > bodyIdx {
		t.Error("expected script to be injected before </body>")
	}
}

func TestInjectAppScript_BeforeHeadClose(t *testing.T) {
	html := []byte(`<html><head><title>Test</title></head><body></body></html>`)
	result := injectAppScript(html)
	s := string(result)
	if !strings.Contains(s, "__heartbeat__") {
		t.Error("expected injected script when only </head> is present")
	}
}

func TestInjectAppScript_AppendedAtEnd(t *testing.T) {
	html := []byte(`<html><h1>No body or head tags</h1></html>`)
	result := injectAppScript(html)
	s := string(result)
	if !strings.Contains(s, "__heartbeat__") {
		t.Error("expected script appended at end when no </body> or </head>")
	}
}

func TestInjectAppScript_AlwaysInject(t *testing.T) {
	// injectAppScript does not check content type; it always injects.
	// Content type detection is handled by htmlInjectWriter.Write().
	html := []byte(`{"key": "value"}`)
	result := injectAppScript(html)
	if !strings.Contains(string(result), "__heartbeat__") {
		t.Error("injectAppScript should always inject regardless of content")
	}
}

func TestActiveRuntimeJS_ProdMode(t *testing.T) {
	devMode = false
	js := activeRuntimeJS()
	if !strings.Contains(js, "contextmenu") {
		t.Error("production JS should disable context menu")
	}
	if !strings.Contains(js, "dragstart") {
		t.Error("production JS should disable drag-and-drop")
	}
}

func TestActiveRuntimeJS_DevMode(t *testing.T) {
	devMode = true
	defer func() { devMode = false }()
	js := activeRuntimeJS()
	if strings.Contains(js, "contextmenu") {
		t.Error("dev mode JS should NOT disable context menu")
	}
	if strings.Contains(js, "dragstart") {
		t.Error("dev mode JS should NOT disable drag-and-drop")
	}
	if !strings.Contains(js, "__heartbeat__") {
		t.Error("dev mode JS should still send heartbeats")
	}
}

func TestBuildSourceContext(t *testing.T) {
	tmpDir := t.TempDir()
	content := "line1\nline2\nline3\nline4\nline5\nline6\nline7\n"
	path := tmpDir + "/test.asp"
	if err := os.WriteFile(path, []byte(content), 0644); err != nil {
		t.Fatal(err)
	}

	html := buildSourceContext(path, 4, 2)
	if html == "" {
		t.Fatal("expected non-empty source context")
	}
	if !strings.Contains(html, `class="errln"`) {
		t.Error("expected error line to be highlighted")
	}
	if !strings.Contains(html, "line4") {
		t.Error("expected line4 content in source context")
	}
}

func TestBuildSourceContext_FileNotFound(t *testing.T) {
	html := buildSourceContext("/nonexistent/file.asp", 1, 3)
	if html != "" {
		t.Error("expected empty string for nonexistent file")
	}
}
