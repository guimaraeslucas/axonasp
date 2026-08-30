package main

import (
	"os"
	"path/filepath"
	"strings"
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

// TestConvertVBScriptTagsToASP_ClientSide verifies that a client-side VBScript
// <script> block (no runat="server") is rewritten to <% %>.
func TestConvertVBScriptTagsToASP_ClientSide(t *testing.T) {
	input := `<html><head>
<script language="VBScript">
MsgBox "Hello"
</script>
</head><body></body></html>`
	result := ConvertVBScriptTagsToASP(input)
	if !strings.Contains(result, `<%`) {
		t.Error("expected <% in result, client-side VBScript block was not converted")
	}
	if strings.Contains(result, `<script language="VBScript">`) {
		t.Error("original <script> tag should be removed from result")
	}
}

// TestConvertVBScriptTagsToASP_RunatServer verifies that runat="server" blocks
// are passed through unchanged (the lexer already handles them).
func TestConvertVBScriptTagsToASP_RunatServer(t *testing.T) {
	input := `<script language="VBScript" runat="server">x = 1</script>`
	result := ConvertVBScriptTagsToASP(input)
	if result != input {
		t.Errorf("runat=server block must pass through unchanged, got %q", result)
	}
}

// TestConvertVBScriptTagsToASP_JavaScript verifies that JavaScript <script>
// blocks are never rewritten.
func TestConvertVBScriptTagsToASP_JavaScript(t *testing.T) {
	input := `<script>alert('hi')</script>`
	result := ConvertVBScriptTagsToASP(input)
	if result != input {
		t.Errorf("plain JS <script> block must pass through unchanged, got %q", result)
	}
}

// TestConvertVBScriptTagsToASP_TypeAttr verifies type="text/vbscript" is also converted.
func TestConvertVBScriptTagsToASP_TypeAttr(t *testing.T) {
	input := `<script type="text/vbscript">Dim x: x = 1</script>`
	result := ConvertVBScriptTagsToASP(input)
	if !strings.Contains(result, `<%`) {
		t.Errorf("type=text/vbscript block should be converted, got %q", result)
	}
}

// TestConvertVBScriptTagsToASP_CaseInsensitive verifies case-insensitive attribute matching.
func TestConvertVBScriptTagsToASP_CaseInsensitive(t *testing.T) {
	input := `<SCRIPT LANGUAGE="VBSCRIPT">Response.Write "x"</SCRIPT>`
	result := ConvertVBScriptTagsToASP(input)
	if !strings.Contains(result, `<%`) {
		t.Errorf("case-insensitive VBScript block should be converted, got %q", result)
	}
}

// TestBuildHTAStyleCSS_BorderMatrix tests all border and borderStyle combinations according to the specification.
func TestBuildHTAStyleCSS_BorderMatrix(t *testing.T) {
	tests := []struct {
		name     string
		cfg      *HtaConfig
		expected string
	}{
		{
			name:     "border=dialog",
			cfg:      &HtaConfig{Border: "dialog"},
			expected: "<style>html, body { border: 3px outset #c0c0c0; box-sizing: border-box; }</style>",
		},
		{
			name:     "border=thin",
			cfg:      &HtaConfig{Border: "thin"},
			expected: "<style>html, body { border: 1px solid #000; box-sizing: border-box; }</style>",
		},
		{
			name:     "borderStyle=normal",
			cfg:      &HtaConfig{BorderStyle: "normal"},
			expected: "",
		},
		{
			name:     "borderStyle=raised",
			cfg:      &HtaConfig{BorderStyle: "raised"},
			expected: "<style>html, body { border: 1px outset #c0c0c0; box-sizing: border-box; }</style>",
		},
		{
			name:     "borderStyle=static",
			cfg:      &HtaConfig{BorderStyle: "static"},
			expected: "<style>html, body { border: 1px solid #000; box-sizing: border-box; }</style>",
		},
		{
			name:     "borderStyle=sunken",
			cfg:      &HtaConfig{BorderStyle: "sunken"},
			expected: "<style>html, body { border: 1px inset #c0c0c0; box-sizing: border-box; }</style>",
		},
		{
			name:     "scroll=yes",
			cfg:      &HtaConfig{Scroll: "yes"},
			expected: "<style>html, body { overflow: scroll !important; }</style>",
		},
		{
			name:     "scroll=no",
			cfg:      &HtaConfig{Scroll: "no"},
			expected: "<style>html, body { overflow: hidden !important; }</style>",
		},
		{
			name:     "scroll=auto",
			cfg:      &HtaConfig{Scroll: "auto"},
			expected: "<style>html, body { overflow: auto !important; }</style>",
		},
		{
			name:     "combined border=dialog and scroll=no",
			cfg:      &HtaConfig{Border: "dialog", Scroll: "no"},
			expected: "<style>html, body { border: 3px outset #c0c0c0; box-sizing: border-box; } html, body { overflow: hidden !important; }</style>",
		},
		{
			name:     "empty config produces empty string",
			cfg:      &HtaConfig{},
			expected: "",
		},
		{
			name:     "nil config produces empty string",
			cfg:      nil,
			expected: "",
		},
	}

	for _, tc := range tests {
		t.Run(tc.name, func(t *testing.T) {
			got := tc.cfg.BuildHTAStyleCSS()
			if got != tc.expected {
				t.Errorf("BuildHTAStyleCSS() = %q, want %q", got, tc.expected)
			}
		})
	}
}

// TestBuildHTAIconHTML tests icon link tag creation.
func TestBuildHTAIconHTML(t *testing.T) {
	cfg := &HtaConfig{Icon: "images/app.ico"}
	expected := `<link rel="icon" href="images/app.ico">`
	if got := cfg.BuildHTAIconHTML(); got != expected {
		t.Errorf("BuildHTAIconHTML() = %q, want %q", got, expected)
	}

	emptyCfg := &HtaConfig{Icon: ""}
	if got := emptyCfg.BuildHTAIconHTML(); got != "" {
		t.Errorf("BuildHTAIconHTML() on empty = %q, want empty string", got)
	}
}

// TestBuildHTAHeadInjections tests combined head injections.
func TestBuildHTAHeadInjections(t *testing.T) {
	cfg := &HtaConfig{
		Icon:   "/favicon.ico",
		Border: "thin",
		Scroll: "no",
	}
	got := cfg.BuildHTAHeadInjections()
	expected := `<link rel="icon" href="/favicon.ico"><style>html, body { border: 1px solid #000; box-sizing: border-box; } html, body { overflow: hidden !important; }</style>`
	if got != expected {
		t.Errorf("BuildHTAHeadInjections() = %q, want %q", got, expected)
	}
}
