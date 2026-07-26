package main

import (
	"os"
	"path/filepath"
	"testing"
)

func TestParseHTATag_FullTag(t *testing.T) {
	dir := t.TempDir()
	path := filepath.Join(dir, "index.hta")
	content := `<html>
<head>
<hta:application applicationname="My App" windowstate="maximize" icon="app.ico" />
</head>
<body>Hello</body>
</html>`
	os.WriteFile(path, []byte(content), 0644)

	cfg := ParseHTATag(path)
	if cfg == nil {
		t.Fatal("expected non-nil config")
	}
	if cfg.ApplicationName != "My App" {
		t.Errorf("expected ApplicationName='My App', got %q", cfg.ApplicationName)
	}
	if cfg.WindowState != "maximize" {
		t.Errorf("expected WindowState='maximize', got %q", cfg.WindowState)
	}
	if cfg.Icon != "app.ico" {
		t.Errorf("expected Icon='app.ico', got %q", cfg.Icon)
	}
}

func TestParseHTATag_NoTag(t *testing.T) {
	dir := t.TempDir()
	path := filepath.Join(dir, "index.html")
	os.WriteFile(path, []byte("<html><body>No HTA tag</body></html>"), 0644)

	cfg := ParseHTATag(path)
	if cfg != nil {
		t.Error("expected nil config when no HTA tag present")
	}
}

func TestParseHTATag_FileNotFound(t *testing.T) {
	cfg := ParseHTATag("/nonexistent/file.hta")
	if cfg != nil {
		t.Error("expected nil config for nonexistent file")
	}
}

func TestStripHTATag(t *testing.T) {
	html := `<html><head><hta:application applicationname="Test" /></head><body>Hi</body></html>`
	result := StripHTATag(html)
	if result != "<html><head></head><body>Hi</body></html>" {
		t.Errorf("unexpected result: %q", result)
	}
}

func TestFindEntryFile(t *testing.T) {
	dir := t.TempDir()

	// No entry file
	result := FindEntryFile(dir)
	if result != "" {
		t.Error("expected empty string when no entry file exists")
	}

	// Create index.hta
	indexPath := filepath.Join(dir, "index.hta")
	os.WriteFile(indexPath, []byte("<html></html>"), 0644)

	result = FindEntryFile(dir)
	if result != indexPath {
		t.Errorf("expected %q, got %q", indexPath, result)
	}
}

func TestFindEntryFile_DefaultAsp(t *testing.T) {
	dir := t.TempDir()

	// Create default.asp (no index.hta)
	defaultPath := filepath.Join(dir, "default.asp")
	os.WriteFile(defaultPath, []byte("<% response.write \"hi\" %>"), 0644)

	result := FindEntryFile(dir)
	if result != defaultPath {
		t.Errorf("expected %q, got %q", defaultPath, result)
	}
}

func TestBoolAttr(t *testing.T) {
	cfg := &HtaConfig{}
	if !cfg.BoolAttr("yes") {
		t.Error("expected true for 'yes'")
	}
	if !cfg.BoolAttr("YES") {
		t.Error("expected true for 'YES'")
	}
	if cfg.BoolAttr("no") {
		t.Error("expected false for 'no'")
	}
	if cfg.BoolAttr("") {
		t.Error("expected false for empty string")
	}
}
